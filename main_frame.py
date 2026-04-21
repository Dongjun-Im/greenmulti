"""초록멀티 메인 프레임"""
import os
import re
import threading
import webbrowser

import requests
import wx
import wx.adv

from config import (
    APP_NAME, APP_VERSION, APP_BUILD_DATE, APP_AUTHOR, APP_EMAIL,
    APP_ADMIN_EMAIL, APP_COPYRIGHT, SORISEM_BASE_URL,
    DATA_DIR,
)
from menu_manager import MenuManager
from page_parser import (
    parse_board_list, parse_post_content, parse_sub_menus,
    PostItem, PostContent, SubMenuItem,
)
from post_dialog import PostDialog
from screen_reader import speak
import theme as theme_mod
from theme import (
    apply_theme, make_font, load_font_size, save_font_size,
    DEFAULT_FONT_SIZE, MIN_FONT_SIZE, MAX_FONT_SIZE, FONT_SIZE_STEP,
    set_current_theme, load_theme_key, get_current_theme_name,
    THEME_PRESETS, THEME_ORDER, init_theme,
)


# 글로벌 다운로드 추적 리스트: [{"name": str, "size": int, "downloaded": int, "status": str}, ...]
download_list = []

# 탐색 상태
VIEW_MAIN_MENU = "main_menu"
VIEW_SUB_MENU = "sub_menu"
VIEW_POST_LIST = "post_list"


class ItemTextCtrl(wx.TextCtrl):
    """특정 키를 완전히 무시하는 TextCtrl (네이티브 레벨 차단)"""
    BLOCKED_KEYS = {
        wx.WXK_HOME, wx.WXK_END, wx.WXK_UP, wx.WXK_DOWN,
        wx.WXK_RETURN, wx.WXK_BACK, wx.WXK_PAGEUP, wx.WXK_PAGEDOWN,
        wx.WXK_ESCAPE,
    }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def MSWHandleMessage(self, msg, wParam, lParam):
        """Windows 메시지 레벨에서 키 차단 (WM_KEYDOWN=0x0100, WM_CHAR=0x0102)"""
        WM_KEYDOWN = 0x0100
        WM_CHAR = 0x0102
        if msg in (WM_KEYDOWN, WM_CHAR):
            # wParam은 virtual key code
            # Home=0x24, End=0x23, Up=0x26, Down=0x28, Return=0x0D,
            # Back=0x08, PageUp=0x21, PageDown=0x22, Escape=0x1B
            # Delete=0x2E 추가
            blocked_vk = {0x24, 0x23, 0x26, 0x28, 0x0D, 0x08, 0x21, 0x22, 0x1B, 0x2E}
            if wParam in blocked_vk:
                return True, 0  # 메시지 처리됨, TextCtrl에 전달하지 않음
        return super().MSWHandleMessage(msg, wParam, lParam)


class GotoDialog(wx.Dialog):
    """바로가기 대화상자 (코드 또는 이름 검색)"""

    def __init__(self, parent, menu_names: list[str],
                 shortcut_codes: list[str] | None = None):
        super().__init__(parent, title="바로가기", style=wx.DEFAULT_DIALOG_STYLE)
        self.all_names = menu_names
        # 번호 코드(앞쪽 숫자) 추출
        self.all_num_codes = [self._extract_num_code(n) for n in menu_names]
        # 바로가기 코드 (URL 기반 문자열 코드)
        self.all_codes = shortcut_codes if shortcut_codes else [""] * len(menu_names)
        self.selected_index = -1
        # 매칭이 없을 때 사용자가 입력한 직접 코드 (bo_table 값으로 추정)
        self.direct_code = ""

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        search_label = wx.StaticText(
            panel, label="바로가기 코드 또는 메뉴 이름 입력(&S):"
        )
        self.search_input = wx.TextCtrl(
            panel, name="바로가기 코드 또는 메뉴 이름", style=wx.TE_PROCESS_ENTER,
        )

        list_label = wx.StaticText(panel, label="메뉴 목록 - 번호. 이름 (바로가기 코드)(&M):")
        self.menu_list = wx.ListBox(
            panel, choices=menu_names, style=wx.LB_SINGLE,
            name="메뉴 목록",
        )
        if menu_names:
            self.menu_list.SetSelection(0)

        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(panel, wx.ID_OK, "이동(&G)")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "취소")
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        ok_btn.SetDefault()

        sizer.Add(search_label, 0, wx.ALL, 5)
        sizer.Add(self.search_input, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
        sizer.Add(list_label, 0, wx.ALL, 5)
        sizer.Add(self.menu_list, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(sizer)
        sizer.Fit(self)
        self.SetMinSize(wx.Size(400, 350))
        self.Fit()

        self.search_input.Bind(wx.EVT_TEXT, self.on_search)
        self.search_input.Bind(wx.EVT_TEXT_ENTER, self.on_enter)
        self.menu_list.Bind(wx.EVT_LISTBOX_DCLICK, self.on_dclick)
        ok_btn.Bind(wx.EVT_BUTTON, self.on_ok)

        # 저시력 테마 적용
        try:
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.search_input.SetFocus()
        self.Centre()

    def _extract_num_code(self, name: str) -> str:
        """메뉴 이름에서 앞쪽 숫자 코드(번호)를 추출한다. 없으면 빈 문자열."""
        m = re.match(r'^(\d+)[\.\)]\s', name)
        return m.group(1) if m else ""

    def _strip_code_suffix(self, name: str) -> str:
        """표시명에서 '(바로가기 코드: X)' 접미사를 제거한다."""
        return re.sub(r'\s*\(바로가기\s*코드:.+?\)\s*$', '', name).strip()

    def on_search(self, event):
        query = self.search_input.GetValue().strip().lower()
        if not query:
            filtered = self.all_names
        elif query.isdigit():
            # 숫자: 번호 정확 매칭 → 번호 접두사 매칭 → 이름 부분일치
            exact = [n for n, c in zip(self.all_names, self.all_num_codes) if c == query]
            if exact:
                filtered = exact
            else:
                filtered = [
                    n for n, c in zip(self.all_names, self.all_num_codes)
                    if c.startswith(query)
                ]
            if not filtered:
                filtered = [
                    n for n in self.all_names
                    if query in self._strip_code_suffix(n).lower()
                ]
        else:
            # 문자: 메뉴 이름(바로가기 코드 접미사 제외)에서만 부분일치 검색
            # 코드 기반 자동 매칭은 하지 않음 → 매칭 없으면 직접 이동으로 자연스럽게 넘어감
            filtered = [
                n for n in self.all_names
                if query in self._strip_code_suffix(n).lower()
            ]
        self.menu_list.Set(filtered)
        if filtered:
            self.menu_list.SetSelection(0)

    def _try_direct_code(self) -> bool:
        """매칭이 없을 때 사용자 입력을 직접 코드로 사용할 수 있는지 확인."""
        query = self.search_input.GetValue().strip()
        if not query:
            return False
        # 영숫자+언더스코어로만 구성된 짧은 코드여야 함
        if not re.match(r'^[a-zA-Z0-9_]{1,40}$', query):
            return False
        self.direct_code = query
        return True

    def on_enter(self, event):
        self._set_selected()
        if self.selected_index >= 0:
            self.EndModal(wx.ID_OK)
            return
        # 매칭된 항목이 없으면 입력한 코드로 직접 이동 시도
        if self._try_direct_code():
            self.EndModal(wx.ID_OK)
            return
        speak("일치하는 메뉴가 없습니다.")

    def on_dclick(self, event):
        self._set_selected()
        if self.selected_index >= 0:
            self.EndModal(wx.ID_OK)

    def on_ok(self, event):
        self._set_selected()
        if self.selected_index >= 0:
            self.EndModal(wx.ID_OK)
            return
        if self._try_direct_code():
            self.EndModal(wx.ID_OK)
            return
        speak("메뉴를 선택하거나 바로가기 코드를 입력해 주세요.")

    def _set_selected(self):
        sel = self.menu_list.GetSelection()
        if sel == wx.NOT_FOUND:
            self.selected_index = -1
            return
        selected_text = self.menu_list.GetString(sel)
        for i, name in enumerate(self.all_names):
            if name == selected_text:
                self.selected_index = i
                return
        self.selected_index = -1

    def get_selection(self) -> int:
        return self.selected_index


class MainFrame(wx.Frame):
    """초록멀티 메인 윈도우"""

    def __init__(self, session: requests.Session):
        super().__init__(
            None,
            title=APP_NAME,
            size=(800, 600),
        )

        self.session = session
        self.menu_manager = MenuManager()
        self.menu_manager.load()

        # 현재 보기 상태
        self.current_view = VIEW_MAIN_MENU
        self.current_sub_menus: list[SubMenuItem] = []
        self.current_posts: list[PostItem] = []
        self.current_menu_name = ""
        self.navigation_stack: list[dict] = []

        # 필드 탐색용 인덱스
        self.field_index = 0
        # 페이지 탐색
        self.current_page = 1
        self.current_board_url = ""
        # 현재 표시 중인 항목 목록 (줄 인덱스 → 데이터 매핑용)
        self.current_items: list[str] = []

        # 저장된 글꼴 크기 불러오기
        self.current_font_size = load_font_size()

        self._build_menu_bar()
        self._build_status_bar()
        self._build_main_panel()
        self._bind_accelerators()

        # 프레임 레벨 키 이벤트
        self.Bind(wx.EVT_CHAR_HOOK, self.on_char_hook)

        # 테마 / 아이콘 적용
        self._apply_full_theme()
        self._set_window_icon()

        self.Centre()
        self.Show()

        if self.current_items:
            speak(f"{APP_NAME}. {self.current_items[0]} 1/{len(self.current_items)}")

    # ── 메뉴바 ──

    def _build_menu_bar(self):
        menubar = wx.MenuBar()

        file_menu = wx.Menu()
        self.id_download_status = wx.NewIdRef()
        self.id_download_dir = wx.NewIdRef()
        self.id_shortcut = wx.NewIdRef()
        self.id_logout = wx.NewIdRef()
        file_menu.Append(self.id_download_status, "다운로드 상태(&J)\tCtrl+J")
        file_menu.AppendSeparator()
        file_menu.Append(self.id_shortcut, "바탕화면 바로가기 만들기(&S)")
        file_menu.AppendSeparator()
        file_menu.Append(self.id_logout, "로그아웃(&L)\tCtrl+L")
        file_menu.AppendSeparator()
        file_menu.Append(wx.ID_EXIT, "프로그램 종료(&X)\tAlt+F4")
        menubar.Append(file_menu, "파일(&F)")

        # 이동 메뉴
        nav_menu = wx.Menu()
        self.id_goto = wx.NewIdRef()
        self.id_goto_main = wx.NewIdRef()
        self.id_goto_page = wx.NewIdRef()
        self.id_search = wx.NewIdRef()
        self.id_page_down = wx.NewIdRef()
        self.id_page_up = wx.NewIdRef()
        self.id_board_refresh = wx.NewIdRef()
        nav_menu.Append(self.id_goto, "바로가기(&G)\tAlt+G")
        nav_menu.Append(self.id_goto_main, "메인 메뉴로 이동(&H)\tAlt+Home")
        nav_menu.AppendSeparator()
        nav_menu.Append(self.id_goto_page, "페이지 이동(&P)\tCtrl+G")
        nav_menu.Append(self.id_page_down, "다음 페이지(&N)\tPageDown")
        nav_menu.Append(self.id_page_up, "이전 페이지(&B)\tPageUp")
        nav_menu.AppendSeparator()
        nav_menu.Append(self.id_search, "게시물 검색(&F)\tCtrl+F")
        nav_menu.Append(self.id_board_refresh, "게시판 새로고침(&R)\tF5")
        menubar.Append(nav_menu, "이동(&N)")

        # 게시물 메뉴
        post_menu = wx.Menu()
        self.id_post_edit = wx.NewIdRef()
        self.id_post_delete = wx.NewIdRef()
        self.id_post_write = wx.NewIdRef()
        post_menu.Append(self.id_post_write, "게시물 작성(&W)\tW")
        post_menu.Append(self.id_post_edit, "게시물 수정(&M)\tAlt+M")
        post_menu.Append(self.id_post_delete, "게시물 삭제(&D)\tAlt+D")
        menubar.Append(post_menu, "게시물(&P)")

        # 설정 메뉴
        settings_menu = wx.Menu()
        self.id_theme = wx.NewIdRef()
        self.id_theme_cycle = wx.NewIdRef()
        self.id_theme_cycle_back = wx.NewIdRef()
        settings_menu.Append(self.id_theme, "화면 테마 선택(&T)\tF7")
        settings_menu.Append(self.id_theme_cycle, "다음 테마로 변경(&N)\tF6")
        settings_menu.Append(self.id_theme_cycle_back, "이전 테마로 변경(&P)\tShift+F6")
        settings_menu.AppendSeparator()
        settings_menu.Append(self.id_download_dir, "다운로드 폴더 변경(&D)")
        menubar.Append(settings_menu, "설정(&S)")

        # 도움말 메뉴
        help_menu = wx.Menu()
        self.id_about = wx.NewIdRef()
        self.id_shortcuts = wx.NewIdRef()
        self.id_mail = wx.NewIdRef()
        help_menu.Append(self.id_about, "프로그램 정보(&A)\tF1")
        help_menu.Append(self.id_shortcuts, "단축키 안내(&K)\tCtrl+K")
        help_menu.AppendSeparator()
        help_menu.Append(self.id_mail, "관리자에게 메일 보내기(&E)\tAlt+E")
        menubar.Append(help_menu, "도움말(&H)")

        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.on_download_status, id=self.id_download_status)
        self.Bind(wx.EVT_MENU, self.on_change_download_dir, id=self.id_download_dir)
        self.Bind(wx.EVT_MENU, self.on_create_shortcut, id=self.id_shortcut)
        self.Bind(wx.EVT_MENU, self.on_logout, id=self.id_logout)
        self.Bind(wx.EVT_MENU, self.on_exit, id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.on_goto, id=self.id_goto)
        self.Bind(wx.EVT_MENU, self._on_menu_goto_main, id=self.id_goto_main)
        self.Bind(wx.EVT_MENU, self._on_menu_goto_page, id=self.id_goto_page)
        self.Bind(wx.EVT_MENU, self._on_menu_search, id=self.id_search)
        self.Bind(wx.EVT_MENU, self._on_menu_page_down, id=self.id_page_down)
        self.Bind(wx.EVT_MENU, self._on_menu_page_up, id=self.id_page_up)
        self.Bind(wx.EVT_MENU, self._on_menu_post_write, id=self.id_post_write)
        self.Bind(wx.EVT_MENU, self._on_menu_post_edit, id=self.id_post_edit)
        self.Bind(wx.EVT_MENU, self._on_menu_post_delete, id=self.id_post_delete)
        self.Bind(wx.EVT_MENU, self.on_about, id=self.id_about)
        self.Bind(wx.EVT_MENU, self.on_shortcuts_help, id=self.id_shortcuts)
        self.Bind(wx.EVT_MENU, self.on_mail, id=self.id_mail)
        self.Bind(wx.EVT_MENU, self.on_change_theme, id=self.id_theme)
        self.Bind(wx.EVT_MENU, self.on_cycle_theme, id=self.id_theme_cycle)
        self.Bind(wx.EVT_MENU, self.on_cycle_theme_back, id=self.id_theme_cycle_back)
        self.Bind(wx.EVT_MENU, self.on_board_refresh, id=self.id_board_refresh)

    def _build_status_bar(self):
        self.status_bar = self.CreateStatusBar(2)
        self.status_bar.SetStatusWidths([-3, -1])
        self.status_bar.SetStatusText("준비", 0)
        self.status_bar.SetStatusText("다운로드 없음", 1)

    def _build_main_panel(self):
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # 현재 항목 1개만 표시하는 TextCtrl
        # 스크린리더가 현재 항목만 읽고 전체를 읽지 않음
        menu_names = self.menu_manager.get_display_names()
        self.current_items = menu_names
        self.current_index = 0

        self.textctrl = ItemTextCtrl(
            self.panel,
            value=menu_names[0] if menu_names else "",
            style=wx.TE_READONLY | wx.TE_DONTWRAP,
            name="메뉴 목록",
        )

        self.sizer.Add(self.textctrl, 0, wx.EXPAND | wx.ALL, 5)
        self.panel.SetSizer(self.sizer)

        self.textctrl.SetInsertionPoint(0)
        self.textctrl.SetFocus()

    def _bind_accelerators(self):
        # 글꼴 확대/축소/원래대로 ID
        self.id_zoom_in = wx.NewIdRef()
        self.id_zoom_out = wx.NewIdRef()
        self.id_zoom_reset = wx.NewIdRef()

        entries = [
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("J"), self.id_download_status),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("K"), self.id_shortcuts),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("L"), self.id_logout),
            wx.AcceleratorEntry(wx.ACCEL_ALT, ord("G"), self.id_goto),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F1, self.id_about),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F5, self.id_board_refresh),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F6, self.id_theme_cycle),
            wx.AcceleratorEntry(wx.ACCEL_SHIFT, wx.WXK_F6, self.id_theme_cycle_back),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F7, self.id_theme),
            wx.AcceleratorEntry(wx.ACCEL_ALT, ord("E"), self.id_mail),
            # 글꼴 크기 단축키
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("="), self.id_zoom_in),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, wx.WXK_NUMPAD_ADD, self.id_zoom_in),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("-"), self.id_zoom_out),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, wx.WXK_NUMPAD_SUBTRACT, self.id_zoom_out),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("0"), self.id_zoom_reset),
        ]
        self.SetAcceleratorTable(wx.AcceleratorTable(entries))

        self.Bind(wx.EVT_MENU, self.on_zoom_in, id=self.id_zoom_in)
        self.Bind(wx.EVT_MENU, self.on_zoom_out, id=self.id_zoom_out)
        self.Bind(wx.EVT_MENU, self.on_zoom_reset, id=self.id_zoom_reset)

    # ── 테마 / 글꼴 ──

    def _apply_full_theme(self):
        """프레임 전체에 테마 색상과 글꼴을 적용한다."""
        font = make_font(self.current_font_size)
        apply_theme(self, font)
        # 상태바는 별도 처리
        try:
            self.status_bar.SetBackgroundColour(theme_mod.COLOR_BG_STATUS)
            self.status_bar.SetForegroundColour(theme_mod.COLOR_FG_STATUS)
            self.status_bar.SetFont(font)
        except Exception:
            pass
        self.Refresh()
        self.Update()

    def _set_window_icon(self):
        """창 아이콘 설정"""
        try:
            icon_path = os.path.join(DATA_DIR, "icon.ico")
            if os.path.exists(icon_path):
                icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
                if icon.IsOk():
                    self.SetIcon(icon)
        except Exception:
            pass

    def on_zoom_in(self, event):
        new_size = min(self.current_font_size + FONT_SIZE_STEP, MAX_FONT_SIZE)
        if new_size == self.current_font_size:
            speak("최대 글꼴 크기입니다.")
            return
        self.current_font_size = new_size
        save_font_size(new_size)
        self._apply_full_theme()
        speak(f"글꼴 크기 {new_size}")

    def on_zoom_out(self, event):
        new_size = max(self.current_font_size - FONT_SIZE_STEP, MIN_FONT_SIZE)
        if new_size == self.current_font_size:
            speak("최소 글꼴 크기입니다.")
            return
        self.current_font_size = new_size
        save_font_size(new_size)
        self._apply_full_theme()
        speak(f"글꼴 크기 {new_size}")

    def on_zoom_reset(self, event):
        self.current_font_size = DEFAULT_FONT_SIZE
        save_font_size(DEFAULT_FONT_SIZE)
        self._apply_full_theme()
        speak(f"글꼴 크기를 원래대로 되돌렸습니다. {DEFAULT_FONT_SIZE}")

    def on_change_theme(self, event):
        """화면 테마 변경 대화상자"""
        current_key = load_theme_key()

        # 테마 목록 구성
        theme_names = []
        for key in THEME_ORDER:
            preset = THEME_PRESETS[key]
            if key == current_key:
                theme_names.append(f"{preset['name']} (현재)")
            else:
                theme_names.append(preset["name"])

        current_idx = THEME_ORDER.index(current_key) if current_key in THEME_ORDER else 0

        dlg = wx.SingleChoiceDialog(
            self,
            "사용할 화면 테마를 선택하세요.\n설정은 자동 저장됩니다.",
            "화면 테마 변경 (F7)",
            theme_names,
        )
        dlg.SetSelection(current_idx)

        # 대화상자에도 테마 적용
        try:
            apply_theme(dlg, make_font(self.current_font_size))
        except Exception:
            pass

        if dlg.ShowModal() == wx.ID_OK:
            sel = dlg.GetSelection()
            new_key = THEME_ORDER[sel]
            if new_key != current_key:
                set_current_theme(new_key)
                self._apply_full_theme()
                new_name = THEME_PRESETS[new_key]["name"]
                speak(f"테마가 {new_name}(으)로 변경되었습니다.")
        dlg.Destroy()

    def on_cycle_theme(self, event):
        """F6: 다음 테마로 순환 변경 (음성 안내 포함)"""
        self._cycle_theme(direction=1)

    def on_cycle_theme_back(self, event):
        """Shift+F6: 이전 테마로 순환 변경 (음성 안내 포함)"""
        self._cycle_theme(direction=-1)

    def _cycle_theme(self, direction: int):
        """테마 순환 변경. direction=+1: 다음, -1: 이전"""
        current_key = load_theme_key()
        try:
            idx = THEME_ORDER.index(current_key)
        except ValueError:
            idx = 0
        next_idx = (idx + direction) % len(THEME_ORDER)
        new_key = THEME_ORDER[next_idx]
        set_current_theme(new_key)
        self._apply_full_theme()
        new_name = THEME_PRESETS[new_key]["name"]
        speak(f"{new_name}")

    # ── 항목 표시 헬퍼 ──

    def _get_current_line_index(self) -> int:
        """현재 선택된 항목의 인덱스를 반환"""
        return self.current_index

    def _format_item(self, line_index: int) -> str:
        """'항목명 N/전체' 형식 문자열 생성"""
        total = len(self.current_items)
        return f"{self.current_items[line_index]} {line_index + 1}/{total}"

    def _move_to_line(self, line_index: int):
        """특정 항목으로 이동 (음성 없이)"""
        if not self.current_items:
            return
        if line_index < 0:
            line_index = 0
        if line_index >= len(self.current_items):
            line_index = len(self.current_items) - 1
        self.current_index = line_index
        # 값만 바꾸고 스크린리더에 맡김
        self.textctrl.ChangeValue(self._format_item(line_index))
        self.textctrl.SetInsertionPoint(0)

    def _jump_to_line_silent(self, line_index: int):
        """Home/End용: NVDA가 글자를 읽는 시점에 TextCtrl을 비워두고,
        전체 텍스트는 speak()로 직접 읽은 뒤 화면 표시를 복원한다."""
        if not self.current_items:
            return
        if line_index < 0:
            line_index = 0
        if line_index >= len(self.current_items):
            line_index = len(self.current_items) - 1
        self.current_index = line_index
        display = self._format_item(line_index)
        # 1. TextCtrl을 비운다 → NVDA가 Home/End 후 읽을 글자가 없음
        self.textctrl.ChangeValue("")
        # 2. 전체 텍스트를 직접 음성 출력
        speak(display)
        # 3. NVDA 키 처리가 끝난 뒤 화면 표시를 복원
        wx.CallLater(80, self._restore_display, display)

    def _restore_display(self, display: str):
        """Home/End 후 화면 표시를 복원한다 (ChangeValue로 스크린리더 무반응)."""
        self.textctrl.ChangeValue(display)
        self.textctrl.SetInsertionPoint(0)

    def _jump_to_line(self, line_index: int):
        """위/아래용: SetValue로 스크린리더가 자동으로 읽도록"""
        if not self.current_items:
            return
        if line_index < 0:
            line_index = 0
        if line_index >= len(self.current_items):
            line_index = len(self.current_items) - 1
        self.current_index = line_index
        display = self._format_item(line_index)
        # SetValue로 스크린리더가 자동으로 이 텍스트를 읽게 함
        self.textctrl.SetValue(display)
        self.textctrl.SetInsertionPoint(0)

    def _update_textctrl(self, items: list[str], label: str):
        """항목 목록을 교체한다."""
        self.current_items = items
        self.current_index = 0
        self.field_index = 0
        self.textctrl.SetName(label)

        if items:
            display = f"{items[0]} 1/{len(items)}"
            self.textctrl.SetValue(display)
            self.textctrl.SetInsertionPoint(0)
            speak(f"{label}. {display}")
        else:
            self.textctrl.SetValue("")
            speak(f"{label}. 항목이 없습니다.")
        self.textctrl.SetFocus()

    # ── 필드 탐색 ──

    def _get_post_fields(self, index: int) -> list[tuple[str, str]]:
        if self.current_view != VIEW_POST_LIST:
            return []
        if index < 0 or index >= len(self.current_posts):
            return []
        post = self.current_posts[index]
        fields = []
        if post.number:
            fields.append(("번호", post.number))
        fields.append(("제목", post.title))
        if post.author:
            fields.append(("작성자", post.author))
        if post.date:
            fields.append(("날짜", post.date))
        if post.comment_count > 0:
            fields.append(("댓글", str(post.comment_count)))
        return fields

    def _read_field(self, direction: int):
        line = self._get_current_line_index()
        fields = self._get_post_fields(line)
        if not fields:
            # 게시글 목록이 아니면 기본 동작
            event_dummy = None
            return False

        self.field_index += direction
        if self.field_index < 0:
            self.field_index = 0
            speak("첫 번째 필드")
            return True
        if self.field_index >= len(fields):
            self.field_index = len(fields) - 1
            speak("마지막 필드")
            return True

        name, value = fields[self.field_index]
        speak(f"{name} '{value}'")
        return True

    # ── 페이지 이동 ──

    def _get_page_url(self, page_num: int) -> str:
        """특정 페이지 번호의 URL을 생성한다."""
        base_url = self.current_board_url
        # 기존 page 파라미터 제거
        base_url = re.sub(r'[?&]page=\d+', '', base_url)
        base_url = base_url.rstrip("&").rstrip("?")
        # 새 page 파라미터 추가
        separator = "&" if "?" in base_url else "?"
        return f"{base_url}{separator}page={page_num}"

    def _navigate_page(self, direction: int):
        if self.current_view != VIEW_POST_LIST:
            speak("게시글 목록에서만 페이지를 이동할 수 있습니다.")
            return
        if not self.current_board_url:
            speak("페이지 이동을 할 수 없습니다.")
            return

        new_page = self.current_page + direction
        if new_page < 1:
            speak("첫 번째 페이지입니다.")
            return

        self.status_bar.SetStatusText(f"{new_page}페이지 로딩 중...", 0)
        speak(f"{new_page}페이지 로딩 중입니다.")

        page_url = self._get_page_url(new_page)

        def on_loaded(html, error):
            if error:
                speak(f"페이지를 불러올 수 없습니다. {error}")
                self.status_bar.SetStatusText("준비", 0)
                return

            if not html or len(html) < 100:
                speak("빈 응답을 받았습니다. 로그인이 필요할 수 있습니다.")
                self.status_bar.SetStatusText("준비", 0)
                return

            posts = parse_board_list(html)
            if posts:
                self.current_page = new_page
                self._show_post_list(posts, self.current_menu_name,
                                     self.current_board_url, new_page)
            else:
                # 디버그: HTML 응답 길이와 로그인 상태 확인
                logged_in = "logout" in html.lower() or "로그아웃" in html
                has_login_form = "login_check" in html and "mb_password" in html
                if has_login_form:
                    speak("로그인이 만료되었습니다. 프로그램을 다시 시작해 주세요.")
                else:
                    speak(f"더 이상 게시글이 없습니다. 현재 {self.current_page}페이지입니다.")
                self.status_bar.SetStatusText("준비", 0)

        self._fetch_page(page_url, on_loaded)

    # ── 화면 표시 ──

    def _show_main_menu(self):
        self.current_view = VIEW_MAIN_MENU
        self.current_sub_menus = []
        self.current_posts = []
        self.current_menu_name = ""
        self.current_board_url = ""
        self.current_page = 1
        self.navigation_stack.clear()
        self.SetTitle(APP_NAME)
        menu_names = self.menu_manager.get_display_names()
        self._update_textctrl(menu_names, "메뉴 목록")
        self.status_bar.SetStatusText("준비", 0)

    def _show_sub_menu(self, sub_menus: list[SubMenuItem], menu_name: str):
        self.current_view = VIEW_SUB_MENU
        self.current_menu_name = menu_name
        self.SetTitle(f"{APP_NAME} - {menu_name}")

        clean_menu = re.sub(r'^\d+[\.\)]\s*', '', menu_name).strip() if menu_name else ""

        # 카테고리 헤더로 제거할 텍스트 목록
        header_noise = {
            "홈", "home",
            "글쓰기", "게시판관리", "멀티업로드",
            "img", "관리자", "철머",
            "로그아웃", "돌아가기",
        }
        # 현재 메뉴명만 노이즈로 추가
        if clean_menu:
            header_noise.add(clean_menu)

        # 브레드크럼(경로) 필터링용: 메인 메뉴 URL 목록
        main_menu_urls = {"/", ""}
        for mi in self.menu_manager.menus:
            main_menu_urls.add(mi.url)

        # 필터링된 하위메뉴와 표시 항목을 동기화
        filtered_subs = []
        display_items = ["0. 메인 메뉴로 돌아가기"]
        seen_texts = set()
        num = 1
        for m in sub_menus:
            # 브레드크럼(경로 안내) 링크 제거: 메인 메뉴 URL과 동일한 항목
            if m.url in main_menu_urls:
                continue

            text = m.display_text

            # 상위 메뉴명 접두사 제거
            if clean_menu and text.startswith(clean_menu):
                text = text[len(clean_menu):].lstrip(" ·:>-")
            elif menu_name and text.startswith(menu_name):
                text = text[len(menu_name):].lstrip(" ·:>-")

            # 기존 번호 제거
            text = re.sub(r'^\d+[\.\)]\s*', '', text).strip()

            # 카테고리 헤더 / 노이즈 제거
            if text.lower() in {h.lower() for h in header_noise}:
                continue
            if not text or len(text) < 2:
                continue
            # 중복 제거
            if text.lower() in seen_texts:
                continue
            seen_texts.add(text.lower())

            # 바로가기 코드 추출 (URL 기반)
            from menu_manager import extract_shortcut_code
            code = extract_shortcut_code(m.url)
            if code:
                display_items.append(f"{num}. {text} (바로가기 코드: {code})")
            else:
                display_items.append(f"{num}. {text}")
            filtered_subs.append(m)
            num += 1

        # 필터링 후 실제 항목이 없으면 "게시물이 없습니다" 표시
        if not filtered_subs:
            display_items = ["0. 메인 메뉴로 돌아가기", "게시물이 없습니다."]

        self.current_sub_menus = filtered_subs
        self._update_textctrl(display_items, f"{menu_name} 하위 메뉴")
        self.status_bar.SetStatusText(f"{menu_name} - {len(filtered_subs)}개 하위 메뉴", 0)

    def _show_post_list(self, posts: list[PostItem], menu_name: str,
                        board_url: str = "", page: int = 1):
        self.current_view = VIEW_POST_LIST
        self.current_posts = posts
        self.current_menu_name = menu_name
        if board_url:
            self.current_board_url = board_url
        self.current_page = page
        self.SetTitle(f"{APP_NAME} - {menu_name}")
        display_items = [p.display_text for p in posts]
        page_label = f" {page}페이지" if page > 1 else ""
        self._update_textctrl(display_items, f"{menu_name}{page_label} 게시글 목록")
        self.status_bar.SetStatusText(
            f"{menu_name} - {page}페이지 ({len(posts)}개 게시글)", 0
        )

    def _show_post_dialog(self, content: PostContent):
        """게시물 대화상자 표시. 이전/다음 게시물 이동도 처리."""
        while True:
            dialog = PostDialog(self, content, self.session)
            result = dialog.ShowModal()
            nav = dialog.navigate_result
            dialog.Destroy()

            if nav == "refresh":
                # 게시물 수정/삭제 후 게시판 새로고침
                if self.current_board_url:
                    self._load_and_show(self.current_board_url, self.current_menu_name)
                break
            elif nav == "prev" and content.prev_url:
                speak("이전 게시물을 불러오는 중입니다.")
                new_content = self._fetch_post_sync(content.prev_url)
                if new_content:
                    content = new_content
                    continue
                else:
                    speak("이전 게시물을 불러올 수 없습니다.")
                    break
            elif nav == "next" and content.next_url:
                speak("다음 게시물을 불러오는 중입니다.")
                new_content = self._fetch_post_sync(content.next_url)
                if new_content:
                    content = new_content
                    continue
                else:
                    speak("다음 게시물을 불러올 수 없습니다.")
                    break
            else:
                break

        self.textctrl.SetFocus()

    def _fetch_post_sync(self, url: str) -> PostContent | None:
        """동기적으로 게시물을 가져와 파싱한다."""
        try:
            full_url = url if url.startswith("http") else f"{SORISEM_BASE_URL}{url}"
            resp = self.session.get(full_url, timeout=15)
            return parse_post_content(resp.text)
        except Exception:
            return None

    # ── 페이지 로딩 ──

    def _fetch_page(self, url: str, callback):
        def worker():
            try:
                full_url = url if url.startswith("http") else f"{SORISEM_BASE_URL}{url}"
                resp = self.session.get(full_url, timeout=15)
                wx.CallAfter(callback, resp.text, None)
            except requests.exceptions.RequestException as e:
                wx.CallAfter(callback, None, str(e))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _load_and_show(self, url: str, name: str):
        self.status_bar.SetStatusText(f"{name} 로딩 중...", 0)
        speak(f"{name} 로딩 중입니다.")

        board_url = url

        def on_loaded(html, error):
            if error:
                speak(f"페이지를 불러올 수 없습니다. {error}")
                self.status_bar.SetStatusText("준비", 0)
                return

            if not html or len(html) < 50:
                speak("빈 응답을 받았습니다.")
                self.status_bar.SetStatusText("준비", 0)
                return

            # 1순위: 게시글 목록
            posts = parse_board_list(html)
            if posts:
                self.current_board_url = board_url
                self._show_post_list(posts, name, board_url, 1)
                return

            # URL에 bo_table이 있으면 게시판 → 글이 0개인 빈 게시판
            if "bo_table=" in board_url:
                self.current_board_url = board_url
                self.current_view = VIEW_POST_LIST
                self.current_posts = []
                self.current_menu_name = name
                self.current_page = 1
                self.SetTitle(f"{APP_NAME} - {name}")
                self._update_textctrl(["게시물이 없습니다."], f"{name} 게시글 목록")
                self.status_bar.SetStatusText(f"{name} - 게시물 없음", 0)
                return

            # 2순위: 하위 메뉴
            sub_menus = parse_sub_menus(html)
            if sub_menus:
                self._show_sub_menu(sub_menus, name)
                return

            # 3순위: 본문
            content = parse_post_content(html)
            if content and content.body:
                self._show_post_dialog(content)
                self.status_bar.SetStatusText("준비", 0)
                return

            # 4순위: 페이지의 모든 의미있는 링크를 하위메뉴로 표시
            from bs4 import BeautifulSoup as _BS
            _soup = _BS(html, "html.parser")

            # script, style, footer, header 영역 제거
            for tag in _soup.find_all(["script", "style", "footer"]):
                tag.decompose()

            fallback_menus = []
            seen = set()
            noise_texts = [
                "본문으로", "상단으로", "로그아웃", "개인정보", "이용약관",
                "돌아가기", "메일", "쪽지", "검색", "홈", "상단", "맨위",
                "저작권", "copyright", "top", "skip",
            ]
            noise_hrefs = [
                "login", "logout", "register", "memo.php", "formmail",
                "mailto:", "password", "javascript:", "history.back",
            ]
            for a in _soup.find_all("a", href=True):
                href = a.get("href", "").strip()
                text = a.get_text(strip=True)
                if not text or len(text) < 2 or len(text) > 60:
                    continue
                if href in ("#", ""):
                    continue
                if any(k in href.lower() for k in noise_hrefs):
                    continue
                if any(k in text for k in noise_texts):
                    continue
                if href.startswith("http") and SORISEM_BASE_URL not in href:
                    # 외부 링크도 포함 (유튜브 등)
                    pass
                elif href.startswith("http"):
                    href = href.replace(SORISEM_BASE_URL, "")

                if href not in seen:
                    seen.add(href)
                    fallback_menus.append(SubMenuItem(text, href))

            if fallback_menus:
                self._show_sub_menu(fallback_menus, name)
                return

            speak(f"{name}에 표시할 내용이 없습니다.")
            self.status_bar.SetStatusText("준비", 0)

        self._fetch_page(url, on_loaded)

    # ── 키보드 이벤트 ──

    def on_char_hook(self, event):
        # TextCtrl에 포커스가 없으면 기본 동작
        if self.FindFocus() != self.textctrl:
            event.Skip()
            return

        keycode = event.GetKeyCode()
        ctrl = event.ControlDown()
        shift = event.ShiftDown()
        alt = event.AltDown()

        if keycode == wx.WXK_RETURN:
            self.on_activate()
        elif keycode == wx.WXK_BACK or keycode == wx.WXK_ESCAPE:
            self.on_go_back()

        # Ctrl+F: 게시물 검색
        elif keycode == ord("F") and ctrl:
            self.on_search()
            return

        # Ctrl+G: 페이지 이동
        elif keycode == ord("G") and ctrl:
            self.on_goto_page()
            return

        # Ctrl+J: 다운로드 상태
        elif keycode == ord("J") and ctrl:
            self.on_download_status(None)
            return

        # Ctrl+K: 단축키 안내
        elif keycode == ord("K") and ctrl:
            self.on_shortcuts_help(None)
            return

        # F1: 프로그램 정보
        elif keycode == wx.WXK_F1:
            self.on_about(None)
            return

        # Alt+Home: 메인 메뉴(초기 화면)으로 돌아가기
        elif keycode == wx.WXK_HOME and alt:
            self._show_main_menu()
            speak("메인 메뉴로 돌아왔습니다.")

        # Home: 첫 항목으로 이동
        elif keycode == wx.WXK_HOME and not alt and not ctrl:
            if self.current_items:
                self._jump_to_line_silent(0)

        # End: 마지막 항목으로 이동
        elif keycode == wx.WXK_END and not alt and not ctrl:
            if self.current_items:
                self._jump_to_line_silent(len(self.current_items) - 1)

        # Shift+좌/우: 필드 읽기
        elif keycode == wx.WXK_LEFT and shift and not ctrl:
            if not self._read_field(-1):
                event.Skip()
        elif keycode == wx.WXK_RIGHT and shift and not ctrl:
            if not self._read_field(1):
                event.Skip()

        # W: 게시물 작성 (게시글 목록에서만)
        elif keycode in (ord("W"), ord("w")) and not ctrl and not alt:
            if self.current_view == VIEW_POST_LIST and self.current_board_url:
                self._write_post()
            else:
                event.Skip()

        # Alt+M: 게시물 수정
        elif keycode in (ord("M"), ord("m")) and alt:
            self._edit_post()

        # Alt+D 또는 Delete: 게시물 삭제
        elif keycode in (ord("D"), ord("d")) and alt:
            self._delete_post()
        elif keycode == wx.WXK_DELETE:
            if self.current_view == VIEW_POST_LIST:
                self._delete_post()

        # PageUp/PageDown: 페이지 이동
        elif keycode == wx.WXK_PAGEUP:
            self._navigate_page(-1)
        elif keycode == wx.WXK_PAGEDOWN:
            self._navigate_page(1)

        # 좌/우, Ctrl+좌/우: TextCtrl 기본 동작 (단일 줄이므로 안전)
        elif keycode in (wx.WXK_LEFT, wx.WXK_RIGHT) and not shift and not alt:
            event.Skip()

        # 위/아래: 줄 이동 + "항목명 N/전체" 읽기
        elif keycode == wx.WXK_UP and not ctrl and not alt:
            self.field_index = 0
            cur = self._get_current_line_index()
            if cur > 0:
                self._jump_to_line(cur - 1)
        elif keycode == wx.WXK_DOWN and not ctrl and not alt:
            self.field_index = 0
            cur = self._get_current_line_index()
            if cur < len(self.current_items) - 1:
                self._jump_to_line(cur + 1)

        else:
            event.Skip()

    # ── 게시물 작성 ──

    def _write_post(self):
        """게시물 작성 대화상자를 표시한다."""
        import re as _re
        # bo_table 추출
        bo_table = ""
        bo_match = _re.search(r'bo_table=([^&]+)', self.current_board_url)
        if bo_match:
            bo_table = bo_match.group(1)

        if not bo_table:
            speak("이 게시판에서는 글을 작성할 수 없습니다.")
            return

        from write_dialog import WriteDialog
        dialog = WriteDialog(self, self.session, bo_table)
        result = dialog.ShowModal()
        dialog.Destroy()

        # 글 작성 성공 시 게시판 새로고침
        if result == wx.ID_OK:
            self._load_and_show(self.current_board_url, self.current_menu_name)

        self.textctrl.SetFocus()

    # ── 게시물 검색 ──

    def on_goto_page(self):
        """페이지 이동 대화상자 (Ctrl+G)"""
        if self.current_view != VIEW_POST_LIST or not self.current_board_url:
            speak("게시글 목록에서만 페이지를 이동할 수 있습니다.")
            return

        dlg = wx.TextEntryDialog(
            self,
            f"현재 {self.current_page}페이지입니다.\n"
            "이동할 페이지 번호를 입력하세요.\n"
            "'끝' 또는 'last'를 입력하면 마지막 페이지로 이동합니다.",
            "페이지 이동",
            str(self.current_page),
        )
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return

        input_text = dlg.GetValue().strip()
        dlg.Destroy()

        if not input_text:
            return

        if input_text in ("끝", "last", "마지막"):
            # 현재 페이지 HTML에서 마지막 페이지 번호 추출
            speak("마지막 페이지를 찾는 중입니다.")
            self._goto_last_page()
            return
        else:
            try:
                target_page = int(input_text)
            except ValueError:
                speak("올바른 페이지 번호를 입력해 주세요.")
                return

        if target_page < 1:
            speak("1 이상의 페이지 번호를 입력해 주세요.")
            return

        if target_page == self.current_page:
            speak(f"이미 {self.current_page}페이지입니다.")
            return

        self._goto_specific_page(target_page)

    def _goto_specific_page(self, target_page: int):
        """특정 페이지로 이동"""
        self.status_bar.SetStatusText(f"{target_page}페이지 로딩 중...", 0)
        speak(f"{target_page}페이지로 이동합니다.")

        page_url = self._get_page_url(target_page)

        def on_loaded(html, error):
            if error:
                speak("페이지를 불러올 수 없습니다.")
                self.status_bar.SetStatusText("준비", 0)
                return

            posts = parse_board_list(html)
            if posts:
                self.current_page = target_page
                self._show_post_list(posts, self.current_menu_name,
                                     self.current_board_url, target_page)
            else:
                speak("해당 페이지에 게시글이 없습니다.")
                self.status_bar.SetStatusText("준비", 0)

        self._fetch_page(page_url, on_loaded)

    def _goto_last_page(self):
        """현재 게시판의 마지막 페이지로 이동"""
        page_url = self._get_page_url(self.current_page)

        def on_loaded(html, error):
            if error:
                speak("페이지를 불러올 수 없습니다.")
                return

            # HTML에서 페이지네이션의 마지막 페이지 번호 추출
            import re as _re
            # page=N 패턴에서 가장 큰 N을 찾기
            page_nums = _re.findall(r'page=(\d+)', html)
            if page_nums:
                last_page = max(int(p) for p in page_nums)
                if last_page > 0:
                    speak(f"마지막 {last_page}페이지로 이동합니다.")
                    self._goto_specific_page(last_page)
                    return

            speak("마지막 페이지를 찾을 수 없습니다.")

        self._fetch_page(page_url, on_loaded)

    def on_search(self):
        """게시물 검색 대화상자"""
        if self.current_view != VIEW_POST_LIST or not self.current_board_url:
            speak("게시글 목록에서만 검색할 수 있습니다.")
            return

        # 검색 유형 → Gnuboard5 sfl 파라미터 매핑
        search_types = [
            ("제목", "wr_subject"),
            ("내용", "wr_content"),
            ("제목+내용", "wr_subject||wr_content"),
            ("회원아이디", "mb_id,1"),
            ("회원아이디(코)", "mb_id,0"),
            ("글쓴이", "wr_name,1"),
        ]

        dlg = wx.Dialog(self, title="게시물 검색", style=wx.DEFAULT_DIALOG_STYLE)
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        type_label = wx.StaticText(panel, label="검색 유형(&T):")
        type_combo = wx.ComboBox(
            panel, choices=[t[0] for t in search_types],
            style=wx.CB_READONLY, name="검색 유형",
        )
        type_combo.SetSelection(0)

        query_label = wx.StaticText(panel, label="검색어(&S):")
        query_input = wx.TextCtrl(panel, name="검색어", style=wx.TE_PROCESS_ENTER)

        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(panel, wx.ID_OK, "검색(&F)")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "취소")
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        ok_btn.SetDefault()

        sizer.Add(type_label, 0, wx.ALL, 5)
        sizer.Add(type_combo, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
        sizer.Add(query_label, 0, wx.ALL, 5)
        sizer.Add(query_input, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        panel.SetSizer(sizer)
        sizer.Fit(dlg)
        dlg.SetMinSize(wx.Size(350, -1))
        dlg.Fit()
        query_input.SetFocus()
        dlg.Centre()

        # Enter키로 검색
        query_input.Bind(wx.EVT_TEXT_ENTER, lambda e: dlg.EndModal(wx.ID_OK))

        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return

        query = query_input.GetValue().strip()
        type_idx = type_combo.GetSelection()
        dlg.Destroy()

        if not query:
            return

        sfl = search_types[type_idx][1]
        type_name = search_types[type_idx][0]
        speak(f"'{query}' {type_name} 검색 중입니다.")

        import urllib.parse
        bo_table = ""
        bo_match = re.search(r'bo_table=([^&]+)', self.current_board_url)
        if bo_match:
            bo_table = bo_match.group(1)

        if not bo_table:
            speak("검색할 수 없는 게시판입니다.")
            return

        search_url = (
            f"/bbs/board.php?bo_table={bo_table}"
            f"&sfl={urllib.parse.quote(sfl)}&stx={urllib.parse.quote(query)}"
        )

        def on_loaded(html, error):
            if error:
                speak("검색에 실패했습니다.")
                return
            posts = parse_board_list(html)
            if posts:
                self.current_board_url = search_url
                self._show_post_list(posts, f"검색: {query}", search_url, 1)
            else:
                speak("검색 결과가 없습니다.")
                wx.MessageBox(
                    f"'{query}' ({type_name}) 검색 결과가 없습니다.",
                    "검색 결과", wx.OK | wx.ICON_INFORMATION, self,
                )

        self._fetch_page(search_url, on_loaded)

    # ── 항목 활성화 ──

    def on_activate(self):
        line = self._get_current_line_index()
        if line < 0 or line >= len(self.current_items):
            return
        if self.current_view == VIEW_MAIN_MENU:
            self._activate_menu(line)
        elif self.current_view == VIEW_SUB_MENU:
            self._activate_sub_menu(line)
        elif self.current_view == VIEW_POST_LIST:
            self._activate_post(line)

    def _activate_menu(self, index: int):
        menu_item = self.menu_manager.get_menu_by_index(index)
        if menu_item is None:
            return

        # 외부 링크: 브라우저에서 열기
        if menu_item.is_external:
            import webbrowser
            speak(f"{menu_item.name}을 브라우저에서 엽니다.")
            webbrowser.open(menu_item.full_url)
            return

        self.navigation_stack.append({
            "view": VIEW_MAIN_MENU,
            "selection": index,
        })
        self._load_and_show(menu_item.url, menu_item.name)

    def _activate_sub_menu(self, index: int):
        # 0번: 메인 메뉴로 돌아가기
        if index == 0:
            self._show_main_menu()
            speak("메인 메뉴로 돌아왔습니다.")
            return
        # 실제 하위 메뉴 인덱스 (0번이 메인메뉴이므로 -1)
        actual = index - 1
        if actual < 0 or actual >= len(self.current_sub_menus):
            return
        sub = self.current_sub_menus[actual]

        # 외부 링크 (http로 시작하고 소리샘이 아닌 URL): 브라우저에서 열기
        full_url = sub.url if sub.url.startswith("http") else f"{SORISEM_BASE_URL}{sub.url}"
        if sub.url.startswith("http") and SORISEM_BASE_URL not in sub.url:
            import webbrowser
            speak(f"{sub.name}을 브라우저에서 엽니다.")
            webbrowser.open(full_url)
            return

        self.navigation_stack.append({
            "view": VIEW_SUB_MENU,
            "sub_menus": self.current_sub_menus,
            "menu_name": self.current_menu_name,
            "selection": index,
        })
        self._load_and_show(sub.url, sub.name)

    def _activate_post(self, index: int):
        if index < 0 or index >= len(self.current_posts):
            return
        post = self.current_posts[index]
        self.status_bar.SetStatusText(f"{post.title} 로딩 중...", 0)
        speak(f"{post.title} 로딩 중입니다.")

        def on_loaded(html, error):
            if error:
                speak(f"게시글을 불러올 수 없습니다. {error}")
                wx.MessageBox(
                    f"게시글을 불러올 수 없습니다.\n{error}",
                    "오류", wx.OK | wx.ICON_ERROR, self,
                )
                self.status_bar.SetStatusText("준비", 0)
                return
            # 디버그: HTML을 임시 파일로 저장 (댓글 구조 분석용)
            import tempfile, os
            debug_path = os.path.join(tempfile.gettempdir(), "chorok_debug.html")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(html)

            content = parse_post_content(html)
            if content:
                self._show_post_dialog(content)
            else:
                speak("게시글 내용을 불러올 수 없습니다.")
                wx.MessageBox(
                    "게시글 내용을 불러올 수 없습니다.",
                    "오류", wx.OK | wx.ICON_ERROR, self,
                )
            self.status_bar.SetStatusText("준비", 0)

        self._fetch_page(post.url, on_loaded)

    def on_go_back(self):
        if not self.navigation_stack:
            if self.current_view != VIEW_MAIN_MENU:
                self._show_main_menu()
            return
        prev = self.navigation_stack.pop()
        sel = prev.get("selection", 0)
        if prev["view"] == VIEW_MAIN_MENU:
            self._show_main_menu()
            self._move_to_line(sel)
        elif prev["view"] == VIEW_SUB_MENU:
            self._show_sub_menu(prev["sub_menus"], prev["menu_name"])
            self._move_to_line(sel)
        elif prev["view"] == VIEW_POST_LIST:
            self.current_board_url = prev.get("board_url", "")
            self.current_page = prev.get("page", 1)
            self._show_post_list(
                prev["posts"], prev["menu_name"],
                self.current_board_url, self.current_page,
            )
            self._move_to_line(sel)

    # ── 메뉴 이벤트 ──

    # ── 메뉴 이벤트 래퍼 ──

    def _on_menu_goto_main(self, event):
        self._show_main_menu()
        speak("메인 메뉴로 돌아왔습니다.")

    def _on_menu_goto_page(self, event):
        self.on_goto_page()

    def _on_menu_search(self, event):
        self.on_search()

    def _on_menu_page_down(self, event):
        self._navigate_page(1)

    def _on_menu_page_up(self, event):
        self._navigate_page(-1)

    def _on_menu_post_write(self, event):
        if self.current_view == VIEW_POST_LIST and self.current_board_url:
            self._write_post()
        else:
            speak("게시글 목록에서만 글을 작성할 수 있습니다.")

    def _on_menu_post_edit(self, event):
        self._edit_post()

    def _on_menu_post_delete(self, event):
        self._delete_post()

    # ── 게시물 bo_table/wr_id 추출 ──

    def _get_post_ids(self, post):
        """게시물 URL 또는 current_board_url에서 bo_table과 wr_id를 추출"""
        bo_table = ""
        wr_id = ""
        # 게시물 URL에서 추출
        url = post.url if hasattr(post, 'url') else ""
        bo_match = re.search(r'bo_table=([^&]+)', url)
        wr_match = re.search(r'wr_id=(\d+)', url)
        if bo_match:
            bo_table = bo_match.group(1)
        if wr_match:
            wr_id = wr_match.group(1)
        # bo_table을 못 찾으면 current_board_url에서 추출
        if not bo_table and self.current_board_url:
            bo_match2 = re.search(r'bo_table=([^&]+)', self.current_board_url)
            if bo_match2:
                bo_table = bo_match2.group(1)
        return bo_table, wr_id

    # ── 게시물 수정 ──

    def _edit_post(self):
        """현재 선택된 게시물 수정"""
        if self.current_view != VIEW_POST_LIST:
            speak("게시글 목록에서만 수정할 수 있습니다.")
            return

        index = self._get_current_line_index()
        if index < 0 or index >= len(self.current_posts):
            return

        post = self.current_posts[index]
        bo_table, wr_id = self._get_post_ids(post)

        if not bo_table or not wr_id:
            speak("이 게시물은 수정할 수 없습니다.")
            return

        # 수정 페이지 URL 생성
        edit_url = f"/bbs/write.php?bo_table={bo_table}&wr_id={wr_id}&w=u"
        speak(f"{post.title} 수정 페이지를 불러오는 중입니다.")

        def on_loaded(html, error):
            if error:
                speak("수정 페이지를 불러올 수 없습니다.")
                return

            # 기존 제목과 본문 추출
            from bs4 import BeautifulSoup as _BS
            soup = _BS(html, "html.parser")

            title_input = soup.find("input", {"name": "wr_subject"})
            body_area = soup.find("textarea", {"name": "wr_content"})

            # 편집 폼이 없으면 권한 없음/에러 → alert 메시지 확인
            if not title_input and not body_area:
                import re as _re
                alert_match = _re.search(r'alert\(["\'](.+?)["\']\)', html)
                if alert_match:
                    speak(f"수정 불가: {alert_match.group(1)}")
                else:
                    speak("이 게시물은 수정할 수 없습니다. 본인이 작성한 글만 수정할 수 있습니다.")
                return

            old_title = title_input.get("value", "") if title_input else post.title
            old_body = body_area.get_text() if body_area else ""

            from write_dialog import WriteDialog
            dialog = WriteDialog(
                self, self.session, bo_table,
                existing_title=old_title, existing_body=old_body,
            )
            # 수정 시 w=u, wr_id 추가
            dialog._edit_wr_id = wr_id

            result = dialog.ShowModal()
            dialog.Destroy()

            if result == wx.ID_OK:
                self._load_and_show(self.current_board_url, self.current_menu_name)

            self.textctrl.SetFocus()

        self._fetch_page(edit_url, on_loaded)

    # ── 게시물 삭제 ──

    def _delete_post(self):
        """현재 선택된 게시물 삭제"""
        if self.current_view != VIEW_POST_LIST:
            speak("게시글 목록에서만 삭제할 수 있습니다.")
            return

        index = self._get_current_line_index()
        if index < 0 or index >= len(self.current_posts):
            return

        post = self.current_posts[index]
        bo_table, wr_id = self._get_post_ids(post)

        if not bo_table or not wr_id:
            speak("이 게시물은 삭제할 수 없습니다.")
            return

        result = wx.MessageBox(
            f"'{post.title}' 게시물을 삭제하시겠습니까?\n\n삭제하면 복구할 수 없습니다.",
            "게시물 삭제", wx.YES_NO | wx.ICON_WARNING, self,
        )
        if result != wx.YES:
            return

        speak("게시물 삭제를 준비하는 중입니다.")

        def worker():
            try:
                import re as _re
                import html as _html

                # 1단계: 게시물 페이지 URL을 정확히 구성
                # post.url에서 bo_table과 wr_id를 추출하여 표준 URL 생성
                post_url = f"{SORISEM_BASE_URL}/bbs/board.php?bo_table={bo_table}&wr_id={wr_id}"
                resp = self.session.get(post_url, timeout=15)

                # "검색어" alert은 URL이 잘못된 것이므로 무시하고 삭제 URL 추출
                # 삭제 URL 추출 (토큰 포함)
                delete_match = _re.search(
                    r'href=["\']([^"\']*delete\.php[^"\']*bo_table=' + _re.escape(bo_table) + r'[^"\']*)["\']',
                    resp.text
                )
                if not delete_match:
                    # 범용 delete.php 검색
                    delete_match = _re.search(
                        r'href=["\']([^"\']*delete\.php[^"\']*)["\']',
                        resp.text
                    )
                if not delete_match:
                    wx.CallAfter(speak, "이 게시물은 삭제할 수 없습니다. 본인이 작성한 글만 삭제 가능합니다.")
                    return

                real_delete_url = _html.unescape(delete_match.group(1))
                if not real_delete_url.startswith("http"):
                    real_delete_url = f"{SORISEM_BASE_URL}{real_delete_url}"

                # 2단계: 토큰 포함 삭제 URL로 실제 삭제 요청
                headers = {"Referer": post_url}
                resp2 = self.session.get(real_delete_url, headers=headers, timeout=15)

                # alert 체크 (삭제 관련 에러만 처리, "검색어" 등 무관한 alert 무시)
                alert_match = _re.search(r'alert\(["\'](.+?)["\']\)', resp2.text)
                if alert_match:
                    msg = alert_match.group(1)
                    if "검색어" in msg:
                        # 삭제와 무관한 alert → 삭제 성공으로 처리
                        wx.CallAfter(self._post_delete_done)
                    else:
                        wx.CallAfter(speak, f"삭제 불가: {msg}")
                        wx.CallAfter(wx.MessageBox,
                                     f"삭제할 수 없습니다.\n{msg}",
                                     "삭제 불가", wx.OK | wx.ICON_WARNING)
                else:
                    wx.CallAfter(self._post_delete_done)
            except Exception as e:
                wx.CallAfter(speak, f"삭제에 실패했습니다. {e}")

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _post_delete_done(self):
        speak("게시물이 삭제되었습니다.")
        wx.MessageBox("게시물이 삭제되었습니다.", "완료",
                      wx.OK | wx.ICON_INFORMATION, self)
        # 게시판 새로고침
        if self.current_board_url:
            self._load_and_show(self.current_board_url, self.current_menu_name)

    def on_board_refresh(self, event):
        """F5: 현재 게시판 목록 새로고침."""
        if self.current_view != VIEW_POST_LIST or not self.current_board_url:
            speak("게시판에서만 새로고침할 수 있습니다.")
            return
        speak("새로고침 중입니다.")
        self._load_and_show(self.current_board_url, self.current_menu_name)

    def on_goto(self, event):
        """바로가기 대화상자: 메인 메뉴 + 현재 하위 메뉴를 모두 포함"""
        from menu_manager import extract_shortcut_code

        # 메인 메뉴 목록
        main_names = self.menu_manager.get_display_names()
        main_codes = self.menu_manager.get_shortcut_codes()

        # 항목: (표시명, 코드, 타입, 인덱스)
        # 표시명은 "N. 이름 (바로가기 코드: xxx)" 형식 유지 (숫자 코드 매칭 위해)
        entries: list[tuple[str, str, str, int]] = []
        for i, (name, code) in enumerate(zip(main_names, main_codes)):
            entries.append((name, code, "main", i))

        # 현재 뷰가 하위 메뉴면 하위 메뉴 항목도 포함
        if self.current_view == VIEW_SUB_MENU and self.current_sub_menus:
            parent = self.current_menu_name or "하위"
            parent_clean = re.sub(r'^\d+[\.\)]\s*', '', parent).strip()
            for i, sm in enumerate(self.current_sub_menus):
                code = extract_shortcut_code(sm.url)
                # 표시명: "N. [부모] 이름 (바로가기 코드: xxx)"
                label = f"{i + 1}. [{parent_clean}] {sm.name}"
                if code:
                    label += f" (바로가기 코드: {code})"
                entries.append((label, code, "sub", i))

        display_names = [e[0] for e in entries]
        codes = [e[1] for e in entries]
        dialog = GotoDialog(self, display_names, codes)
        if dialog.ShowModal() == wx.ID_OK:
            sel = dialog.get_selection()
            direct_code = getattr(dialog, "direct_code", "") or ""
            if 0 <= sel < len(entries):
                _, _, kind, idx = entries[sel]
                if kind == "main":
                    self._show_main_menu()
                    self._move_to_line(idx)
                    self._activate_menu(idx)
                elif kind == "sub":
                    self._move_to_line(idx + 1)
                    self._activate_sub_menu(idx + 1)
            elif direct_code:
                # 매칭이 없을 때 사용자가 입력한 코드로 직접 이동
                # 1) 클럽(cl=) → 2) 게시판(bo_table=) 순서로 시도
                self._navigate_by_direct_code(direct_code)
        dialog.Destroy()

    def _extract_page_title(self, html: str, fallback: str,
                            match_code: str = "") -> str:
        """HTML에서 페이지 제목(클럽/게시판 이름)을 추출한다.

        match_code가 주어지면 그 코드(cl=xxx, bo_table=xxx)를 가리키는
        링크의 텍스트를 우선 사용한다.
        """
        from bs4 import BeautifulSoup as _BS
        soup = _BS(html, "html.parser")

        bad_words = {
            "sorisem", "소리샘", "home", "홈", "메뉴", "로그인",
            "안내", "관리", "검색", "회원", "main", "index",
            "body", "content", "페이지", "목록", "게시판",
            "공지사항", "자유게시판", "자료실", "자유",
        }
        fallback_lower = fallback.lower() if fallback else ""

        def _is_valid(t: str) -> bool:
            if not t or len(t) < 2 or len(t) > 80:
                return False
            if t.lower() in bad_words:
                return False
            if fallback_lower and t.lower() == fallback_lower:
                return False
            if not re.search(r'[가-힣a-zA-Z]', t):
                return False
            # "N. X" 또는 "N) X" 형식은 하위메뉴/게시판 항목 - 클럽 제목이 아님
            if re.match(r'^\d+[\.\)]\s', t):
                return False
            return True

        # ⭐ 최우선: 현재 코드를 가리키는 링크의 텍스트
        # 예: <a href="/plugin/ar.club/?cl=hims">셀바스헬스케어(구) 힘스인터네셔널</a>
        # 주의: cl=hims 가 포함된 하위 게시판 URL(bo_table=xxx&cl=hims)은 제외
        if match_code:
            candidates = []
            for a in soup.find_all("a", href=True):
                href = a.get("href", "")
                # 1) 클럽 메인 페이지: cl=CODE 포함, bo_table 없음
                is_club_main = (
                    f"cl={match_code}" in href
                    and "bo_table=" not in href
                )
                # 2) 게시판 메인 페이지: bo_table=CODE
                is_board_main = f"bo_table={match_code}" in href
                if not (is_club_main or is_board_main):
                    continue
                t = a.get_text(" ", strip=True)
                if not t:
                    t = a.get("title", "").strip()
                if _is_valid(t):
                    candidates.append(t)
            if candidates:
                # 가장 긴(구체적인) 이름 선택
                return max(candidates, key=len)

        # 브레드크럼 탐색
        for sel in [
            ".breadcrumb", "#breadcrumb",
            ".bread", ".crumb", "nav.crumb",
            ".location", "#location",
            ".path", "#nav_path",
        ]:
            el = soup.select_one(sel)
            if el:
                crumbs = [a.get_text(strip=True) for a in el.find_all(["a", "span", "li"])]
                crumbs = [c for c in crumbs if _is_valid(c)]
                if crumbs:
                    return crumbs[-1]

        # 활성화된 네비게이션 항목
        for sel in [
            "#gnb li.current", "#gnb li.active", "#gnb li.on",
            "#lnb li.current", "#lnb li.active", "#lnb li.on",
            "#snb li.current", "#snb li.active", "#snb li.on",
            ".gnb li.current", ".gnb li.active", ".gnb li.on",
            ".menu li.current", ".menu li.active", ".menu li.on",
            ".nav li.current", ".nav li.active", ".nav li.on",
            "a.current", "a.active", "a.on",
        ]:
            for el in soup.select(sel)[:5]:
                t = el.get_text(strip=True)
                if _is_valid(t):
                    return t

        # 1. 소리샘/Gnuboard ar.club 플러그인 특화 셀렉터
        for sel in [
            "#bo_cate_on", ".bo_cate_on",
            ".bo_cate h1", ".bo_cate h2", ".bo_cate strong",
            "#bo_gall_cate strong", "#bo_gall_cate",
            ".ar_club_title", ".ar_title", ".club_name",
            ".club_title", "#club_title",
            ".board_title", "#board_title",
            "h1.title", "h2.title", "h3.title",
            ".write_info .subject",
            ".page_title", ".sub_title",
            "article h1", "article h2",
            "#container h1", "#container h2",
            "#wrapper h1", "#wrapper h2",
            "main h1", "main h2",
            ".content h1", ".content h2",
        ]:
            el = soup.select_one(sel)
            if el:
                t = el.get_text(strip=True)
                if _is_valid(t):
                    return t

        # 2. 활성화된 메뉴/카테고리 마커
        for sel in [".active", ".current", ".on", ".selected"]:
            for el in soup.select(sel):
                t = el.get_text(strip=True)
                if _is_valid(t):
                    return t

        # 3. <h1>, <h2>, <h3> 순서로 첫 번째 유효한 것
        for tag in ["h1", "h2", "h3"]:
            for el in soup.find_all(tag):
                t = el.get_text(strip=True)
                if _is_valid(t):
                    return t

        # 4. <title> 태그: 여러 구분자로 분리
        title_el = soup.find("title")
        if title_el:
            t = title_el.get_text(strip=True)
            parts = re.split(r'\s*[-|:：>·•/\|]+\s*', t)
            parts = [p.strip() for p in parts if p.strip()]
            meaningful = [p for p in parts if _is_valid(p)]
            if meaningful:
                # 가장 긴(구체적인) 부분
                return max(meaningful, key=len)

        # 5. og:title 메타 태그
        for meta_key in [
            ("property", "og:title"),
            ("name", "title"),
            ("name", "og:title"),
        ]:
            meta = soup.find("meta", {meta_key[0]: meta_key[1]})
            if meta:
                t = meta.get("content", "").strip()
                if _is_valid(t):
                    return t

        # 6. body의 strong/b 태그 중 가장 앞쪽 유효한 것
        for el in soup.find_all(["strong", "b"])[:10]:
            t = el.get_text(strip=True)
            if _is_valid(t):
                return t

        return fallback

    def _navigate_by_direct_code(self, code: str):
        """사용자 입력 코드를 클럽 → 게시판 순서로 시도하여 이동한다."""
        club_url = f"/plugin/ar.club/?cl={code}"
        board_url = f"/bbs/board.php?bo_table={code}"
        # 간결한 안내: _load_and_show와 동일 패턴
        self.status_bar.SetStatusText(f"{code} 로딩 중...", 0)
        speak(f"{code} 로딩 중입니다.")

        def resolve_display_name(html, sub_menus) -> str:
            """페이지 제목 추출. 실패 시 사용자가 입력한 코드 사용."""
            display_name = self._extract_page_title(html, "", match_code=code)
            if display_name:
                return display_name
            # 하위메뉴 첫 항목 폴백은 엉뚱한 결과(공지사항, FAQ 등)를
            # 반환하여 혼란을 주므로 사용하지 않음 - 그냥 코드 반환
            return code

        def render_from_html(html, tried_url, attempted_board: bool) -> bool:
            if not html or len(html) < 100:
                return False
            posts = parse_board_list(html)
            sub_menus = parse_sub_menus(html)
            display_name = resolve_display_name(html, sub_menus)
            # 클럽 URL (ar.club 플러그인)이면 하위메뉴 우선
            is_club_url = "ar.club" in tried_url
            if is_club_url and sub_menus:
                self._show_sub_menu(sub_menus, display_name)
                return True
            if posts:
                self.current_board_url = tried_url
                self._show_post_list(posts, display_name, tried_url, 1)
                return True
            if sub_menus:
                self._show_sub_menu(sub_menus, display_name)
                return True
            if "bo_table=" in tried_url and attempted_board:
                self.current_board_url = tried_url
                self.current_view = VIEW_POST_LIST
                self.current_posts = []
                self.current_menu_name = display_name
                self.current_page = 1
                self.SetTitle(f"{APP_NAME} - {display_name}")
                self._update_textctrl(
                    ["게시물이 없습니다."], f"{display_name} 게시글 목록"
                )
                self.status_bar.SetStatusText(f"{display_name} - 게시물 없음", 0)
                return True
            return False

        def on_board_loaded(html, error):
            if error or not render_from_html(html, board_url, attempted_board=True):
                speak("표시할 내용이 없습니다.")
                self.status_bar.SetStatusText("준비", 0)

        def on_club_loaded(html, error):
            if not error and render_from_html(html, club_url, attempted_board=False):
                return
            self._fetch_page(board_url, on_board_loaded)

        self._fetch_page(club_url, on_club_loaded)

    # ── 다운로드 상태 (Ctrl+J) ──

    def on_download_status(self, event):
        """다운로드 상태 대화상자"""
        from config import get_download_dir

        dlg = wx.Dialog(self, title="파일 다운로드", size=(500, 400),
                        style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        dl_dir = get_download_dir()
        dir_label = wx.StaticText(panel, label=f"다운로드 폴더: {dl_dir}")
        sizer.Add(dir_label, 0, wx.ALL, 5)

        label = wx.StaticText(panel, label="다운로드 목록(&L):")
        dl_list_ctrl = wx.ListBox(panel, name="다운로드 목록")

        # 다운로드 상태 표시
        items = []
        for dl in download_list:
            name = dl["name"]
            status = dl["status"]
            size = dl["size"]
            downloaded = dl["downloaded"]

            if size > 0:
                if size < 1024:
                    size_str = f"{size}B"
                elif size < 1024 * 1024:
                    size_str = f"{size / 1024:.1f}KB"
                else:
                    size_str = f"{size / (1024 * 1024):.1f}MB"
                pct = int(downloaded / size * 100) if size > 0 else 0
            else:
                size_str = "알 수 없음"
                pct = 0

            if status == "완료":
                items.append(f"{name} ({size_str}) - 완료")
            elif status == "실패":
                items.append(f"{name} - 실패")
            else:
                items.append(f"{name} ({size_str}) - {pct}%")

        dl_list_ctrl.Set(items if items else ["다운로드한 파일이 없습니다."])

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        delete_btn = wx.Button(panel, label="파일 삭제(&D)")
        clear_btn = wx.Button(panel, label="목록 비우기(&C)")
        close_btn = wx.Button(panel, wx.ID_CANCEL, "닫기(&X)")
        btn_sizer.Add(delete_btn, 0, wx.RIGHT, 5)
        btn_sizer.Add(clear_btn, 0, wx.RIGHT, 5)
        btn_sizer.AddStretchSpacer()
        btn_sizer.Add(close_btn, 0)

        sizer.Add(label, 0, wx.LEFT | wx.RIGHT, 5)
        sizer.Add(dl_list_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 5)
        panel.SetSizer(sizer)
        dl_list_ctrl.SetFocus()

        def on_delete(evt):
            sel = dl_list_ctrl.GetSelection()
            if sel == wx.NOT_FOUND or sel >= len(download_list):
                speak("삭제할 파일을 선택해 주세요.")
                return
            dl_entry = download_list[sel]
            fname = dl_entry["name"]
            fpath = os.path.join(dl_dir, fname)
            r = wx.MessageBox(
                f"'{fname}' 파일을 삭제하시겠습니까?",
                "파일 삭제", wx.YES_NO | wx.ICON_QUESTION, dlg,
            )
            if r == wx.YES:
                try:
                    if os.path.exists(fpath):
                        os.remove(fpath)
                    download_list.pop(sel)
                    # 목록 갱신
                    new_items = []
                    for dl in download_list:
                        n = dl["name"]
                        s = dl["status"]
                        sz = dl["size"]
                        if s == "완료":
                            if sz < 1024:
                                ss = f"{sz}B"
                            elif sz < 1024*1024:
                                ss = f"{sz/1024:.1f}KB"
                            else:
                                ss = f"{sz/(1024*1024):.1f}MB"
                            new_items.append(f"{n} ({ss}) - 완료")
                        elif s == "실패":
                            new_items.append(f"{n} - 실패")
                        else:
                            new_items.append(f"{n} - 다운로드 중")
                    dl_list_ctrl.Set(new_items if new_items else ["다운로드한 파일이 없습니다."])
                    speak(f"{fname} 파일이 삭제되었습니다.")
                except Exception as e:
                    speak(f"삭제에 실패했습니다. {e}")

        def on_clear(evt):
            download_list.clear()
            dl_list_ctrl.Set(["다운로드한 파일이 없습니다."])
            speak("다운로드 목록을 비웠습니다.")

        delete_btn.Bind(wx.EVT_BUTTON, on_delete)
        clear_btn.Bind(wx.EVT_BUTTON, on_clear)

        dlg.Centre()
        dlg.ShowModal()
        dlg.Destroy()

    # ── 다운로드 폴더 변경 ──

    def on_change_download_dir(self, event):
        """다운로드 폴더 변경"""
        from config import get_download_dir, set_download_dir
        current = get_download_dir()
        dlg = wx.DirDialog(
            self, "다운로드 폴더를 선택하세요",
            defaultPath=current,
            style=wx.DD_DEFAULT_STYLE,
        )
        if dlg.ShowModal() == wx.ID_OK:
            new_dir = dlg.GetPath()
            set_download_dir(new_dir)
            speak(f"다운로드 폴더가 변경되었습니다. {new_dir}")
            wx.MessageBox(f"다운로드 폴더가 변경되었습니다.\n{new_dir}",
                          "완료", wx.OK | wx.ICON_INFORMATION, self)
        dlg.Destroy()

    # ── 바탕화면 바로가기 ──

    def on_create_shortcut(self, event):
        """바탕화면에 바로가기 만들기"""
        try:
            import sys
            import ctypes
            from ctypes import wintypes

            # 바탕화면 경로: 로컬 바탕화면 (OneDrive가 아닌 실제 경로)
            # 1순위: C:\Users\사용자\Desktop
            user_profile = os.environ.get("USERPROFILE", os.path.expanduser("~"))
            desktop = os.path.join(user_profile, "Desktop")
            if not os.path.exists(desktop):
                desktop = os.path.join(user_profile, "바탕 화면")
            if not os.path.exists(desktop):
                # 2순위: PUBLIC Desktop
                public = os.environ.get("PUBLIC", "")
                if public:
                    desktop = os.path.join(public, "Desktop")
            if not os.path.exists(desktop):
                speak("바탕화면 폴더를 찾을 수 없습니다.")
                return

            if getattr(sys, 'frozen', False):
                target = sys.executable
            else:
                target = os.path.abspath(sys.argv[0])

            shortcut_path = os.path.join(desktop, f"{APP_NAME}.lnk")

            # 아이콘 경로 (초록 나무 아이콘)
            icon_path = os.path.join(DATA_DIR, "icon.ico")

            # win32com으로 바로가기 생성
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            sc = shell.CreateShortCut(shortcut_path)
            sc.TargetPath = target
            sc.WorkingDirectory = os.path.dirname(target)
            sc.Description = APP_NAME
            if os.path.exists(icon_path):
                sc.IconLocation = f"{icon_path},0"
            sc.save()

            speak("바탕화면에 바로가기를 만들었습니다.")
            wx.MessageBox(
                f"바탕화면에 바로가기를 만들었습니다.\n{shortcut_path}",
                "완료", wx.OK | wx.ICON_INFORMATION, self,
            )

        except Exception as e:
            speak("바로가기 만들기에 실패했습니다.")
            wx.MessageBox(
                f"바로가기 만들기에 실패했습니다.\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )
            wx.MessageBox(f"바로가기 만들기에 실패했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    def on_logout(self, event):
        result = wx.MessageBox(
            "로그아웃하면 프로그램이 종료됩니다.\n로그아웃하시겠습니까?",
            "로그아웃", wx.YES_NO | wx.ICON_QUESTION, self,
        )
        if result == wx.YES:
            from green_auth import delete_credentials
            delete_credentials()
            try:
                from green_auth.config import LOGOUT_URL
                self.session.get(LOGOUT_URL, timeout=10)
            except Exception:
                pass
            speak("로그아웃되었습니다.")
            self.Close()

    # ── 프로그램 정보 (F1) ──

    def on_about(self, event):
        info = wx.adv.AboutDialogInfo()
        info.SetName(APP_NAME)
        info.SetVersion(f"{APP_VERSION} (빌드 날짜: {APP_BUILD_DATE})")
        info.SetDescription(
            "소리샘 시각장애인 사이트 전용 프로그램\n\n"
            f"제작자: {APP_AUTHOR}\n"
            f"이메일: {APP_EMAIL}\n"
            f"빌드 날짜: {APP_BUILD_DATE}"
        )
        info.SetCopyright(APP_COPYRIGHT)
        info.AddDeveloper(f"{APP_AUTHOR} ({APP_EMAIL})")
        wx.adv.AboutBox(info, self)

    # ── 단축키 안내 (Ctrl+K) ──

    def on_shortcuts_help(self, event):
        shortcuts = [
            "=== 메뉴 탐색 ===",
            "위/아래 방향키: 항목 이동",
            "좌/우 방향키: 글자 단위 읽기",
            "Ctrl+좌/우: 단어 단위 읽기",
            "Shift+좌/우: 필드 단위 읽기 (번호/제목/작성자/날짜)",
            "Enter: 항목 선택/진입",
            "Backspace/ESC: 뒤로 가기",
            "Home: 첫 항목으로 이동",
            "End: 마지막 항목으로 이동",
            "",
            "=== 이동 ===",
            "Alt+Home: 메인 메뉴로 이동",
            "Alt+G: 바로가기",
            "Ctrl+G: 페이지 이동",
            "PageDown: 다음 페이지",
            "PageUp: 이전 페이지",
            "Ctrl+F: 게시물 검색",
            "F5: 게시판 새로고침",
            "",
            "=== 게시물 ===",
            "W: 게시물 작성",
            "Alt+M: 게시물 수정",
            "Alt+D / Delete: 게시물 삭제",
            "Alt+R: 게시물 답변",
            "",
            "=== 게시물 본문 ===",
            "Enter: 커서 위치의 URL을 브라우저에서 열기",
            "Ctrl+U: 게시물 내 URL 목록 보기",
            "B: 본문 txt 저장",
            "Alt+S: 첨부파일 저장",
            "Alt+M: 게시물 수정",
            "Alt+D: 게시물 삭제",
            "Alt+R: 게시물 답변",
            "C: 댓글 작성",
            "D: 댓글 삭제 (댓글 목록에서)",
            "M: 댓글 수정 (댓글 목록에서)",
            "N: 댓글 정렬 순서 변경",
            "Alt+B: 이전 게시물",
            "Alt+N: 다음 게시물",
            "",
            "=== 기타 ===",
            "Ctrl+J: 다운로드 상태",
            "Ctrl+K: 단축키 안내",
            "Ctrl+L: 로그아웃",
            "Alt+E: 관리자에게 메일 보내기",
            "F1: 프로그램 정보",
            "Alt+F4: 프로그램 종료",
            "",
            "=== 화면 설정 (저시력 지원) ===",
            "F7: 화면 테마 선택 (목록)",
            "F6: 다음 테마로 변경",
            "Shift+F6: 이전 테마로 변경",
            "Ctrl++: 글꼴 크게 (확대)",
            "Ctrl+-: 글꼴 작게 (축소)",
            "Ctrl+0: 글꼴 크기 원래대로",
        ]

        dlg = wx.Dialog(self, title="단축키 안내", size=(450, 500),
                        style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)
        text = wx.TextCtrl(
            panel, value="\n".join(shortcuts),
            style=wx.TE_MULTILINE | wx.TE_READONLY,
            name="단축키 목록",
        )
        close_btn = wx.Button(panel, wx.ID_CANCEL, "닫기(&X)")
        sizer.Add(text, 1, wx.EXPAND | wx.ALL, 10)
        sizer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(sizer)
        text.SetFocus()
        text.SetInsertionPoint(0)
        dlg.Centre()
        dlg.ShowModal()
        dlg.Destroy()

    def on_mail(self, event):
        webbrowser.open(f"mailto:{APP_ADMIN_EMAIL}")

    def on_exit(self, event):
        self.Close()
