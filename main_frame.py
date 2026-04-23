"""초록멀티 메인 프레임"""
import os
import re
import sys
import threading
import webbrowser

import requests
import wx
import wx.adv

from config import (
    APP_NAME, APP_VERSION, APP_BUILD_DATE, APP_AUTHOR, APP_EMAIL,
    APP_ADMIN_EMAIL, APP_COPYRIGHT, SORISEM_BASE_URL,
    DATA_DIR,
    load_search_history, add_search_history,
    resource_path,
    load_update_settings, save_update_settings, UPDATE_RELEASES_PAGE,
    get_update_interval_hours,
)
from updater import (
    check_latest_release, is_newer, ReleaseInfo,
    download_installer, get_download_dir, DownloadCancelled,
    ChecksumMismatch, sha256_of_file, fetch_expected_checksum,
    clean_release_notes,
    detect_installation_kind, get_install_dir, fetch_manifest, compute_delta,
    extract_zip, write_restart_script,
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

        # UI가 뜬 뒤 NAS 자동 마운트 시도 (저장된 자격증명이 있을 때만, 백그라운드)
        wx.CallLater(500, self._try_auto_mount_nas)

        # 시작 시 자동 업데이트 확인 (설정에서 끌 수 있음). 로그인/메뉴 음성이
        # 먼저 끝나도록 몇 초 지연 후 백그라운드로 실행.
        wx.CallLater(3000, self._auto_update_check)

        # 쪽지 실시간 알림 폴링 시작 (1분 간격)
        self._unread_memo_count = 0
        self._unread_mail_count = 0
        self._base_title = APP_NAME
        wx.CallLater(5000, self._start_memo_notifier)

    def _try_auto_mount_nas(self):
        """저장된 NAS 자격증명이 있으면 백그라운드로 rclone 마운트 시도."""
        try:
            from nas import (
                load_nas_credentials, mount, find_existing_mount,
                _is_winfsp_missing, WINFSP_DOWNLOAD_URL,
            )
        except Exception:
            return

        existing = find_existing_mount()
        if existing:
            speak("초록등대 자료실에 연결되었습니다.")
            return

        creds = load_nas_credentials()
        if not creds:
            return
        user, pw = creds

        import threading

        def worker():
            ok, info = mount(user, pw)
            if ok:
                wx.CallAfter(self._notify_nas_connected)
                return
            if _is_winfsp_missing(info):
                wx.CallAfter(self._prompt_winfsp_install, WINFSP_DOWNLOAD_URL)
                return
            wx.CallAfter(speak, "초록등대 자료실 자동 연결에 실패했습니다.")
            wx.CallAfter(
                wx.MessageBox,
                f"초록등대 자료실 자동 연결에 실패했습니다.\n\n{info}",
                "NAS 자동 연결 실패", wx.OK | wx.ICON_ERROR, self,
            )

        threading.Thread(target=worker, daemon=True).start()

    def _notify_nas_connected(self):
        """연결 성공 시 음성 안내 + 정보 팝업."""
        speak("초록등대 자료실에 연결되었습니다.")
        wx.MessageBox(
            "초록등대 자료실에 연결되었습니다.",
            "연결 완료", wx.OK | wx.ICON_INFORMATION, self,
        )

    def _prompt_winfsp_install(self, download_url: str):
        """WinFSP 미설치 안내. 링크를 브라우저로 열어 줌."""
        r = wx.MessageBox(
            "드라이브 문자 매핑에 필요한 WinFSP 가 이 PC에 설치되어 있지 않습니다.\n\n"
            "WinFSP 는 오픈소스 Windows 파일 시스템 드라이버로, 초록등대 자료실을 "
            "일반 드라이브처럼 쓰기 위해 한 번만 설치하면 됩니다.\n\n"
            "지금 다운로드 페이지를 열까요? (페이지에서 최신 .msi 파일을 받아 설치하세요)",
            "WinFSP 설치 필요",
            wx.YES_NO | wx.ICON_INFORMATION, self,
        )
        if r == wx.YES:
            try:
                os.startfile(download_url)
            except Exception as e:
                wx.MessageBox(
                    f"브라우저 열기 실패.\n\n{download_url}\n\n{e}",
                    "오류", wx.OK | wx.ICON_ERROR, self,
                )

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
        self.id_settings = wx.NewIdRef()
        settings_menu.Append(self.id_settings, "설정(&T)\tF7")
        settings_menu.AppendSeparator()
        settings_menu.Append(self.id_download_dir, "다운로드 폴더 변경(&D)")
        menubar.Append(settings_menu, "설정(&S)")

        # 도구 메뉴
        tools_menu = wx.Menu()
        self.id_nas_connect = wx.NewIdRef()
        self.id_memo_inbox = wx.NewIdRef()
        self.id_memo_compose = wx.NewIdRef()
        self.id_mail_compose = wx.NewIdRef()
        tools_menu.Append(self.id_nas_connect, "초록등대 자료실 연결(&N)\tCtrl+N")
        tools_menu.AppendSeparator()
        tools_menu.Append(self.id_memo_inbox, "쪽지함 열기(&M)\tCtrl+M")
        tools_menu.Append(self.id_memo_compose, "쪽지 쓰기\tCtrl+Shift+M")
        tools_menu.Append(self.id_mail_compose, "메일함 열기\tCtrl+Shift+E")
        self.id_memo_check_now = wx.NewIdRef()
        tools_menu.Append(self.id_memo_check_now, "알림 센터 열기\tCtrl+Shift+N")
        menubar.Append(tools_menu, "도구(&T)")

        # 도움말 메뉴
        help_menu = wx.Menu()
        self.id_about = wx.NewIdRef()
        self.id_shortcuts = wx.NewIdRef()
        self.id_mail = wx.NewIdRef()
        self.id_manual = wx.NewIdRef()
        self.id_update_check = wx.NewIdRef()
        help_menu.Append(self.id_about, "프로그램 정보(&A)\tF1")
        help_menu.Append(self.id_manual, "사용자 설명서(&U)\tShift+F1")
        help_menu.Append(self.id_shortcuts, "단축키 안내(&K)\tCtrl+K")
        help_menu.AppendSeparator()
        help_menu.Append(self.id_update_check, "업데이트 확인(&P)\tAlt+U")
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
        self.Bind(wx.EVT_MENU, self.on_show_manual, id=self.id_manual)
        self.Bind(wx.EVT_MENU, self.on_shortcuts_help, id=self.id_shortcuts)
        self.Bind(wx.EVT_MENU, self.on_mail, id=self.id_mail)
        self.Bind(wx.EVT_MENU, self.on_manual_update_check, id=self.id_update_check)
        self.Bind(wx.EVT_MENU, self.on_show_settings, id=self.id_settings)
        self.Bind(wx.EVT_MENU, self.on_board_refresh, id=self.id_board_refresh)
        self.Bind(wx.EVT_MENU, self._on_menu_nas_connect, id=self.id_nas_connect)
        self.Bind(wx.EVT_MENU, self.on_open_memo_inbox, id=self.id_memo_inbox)
        self.Bind(wx.EVT_MENU, self.on_open_memo_compose, id=self.id_memo_compose)
        self.Bind(wx.EVT_MENU, self.on_open_mail_compose, id=self.id_mail_compose)
        self.Bind(wx.EVT_MENU, self.on_memo_check_now, id=self.id_memo_check_now)

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
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("N"), self.id_nas_connect),
            wx.AcceleratorEntry(wx.ACCEL_CTRL, ord("M"), self.id_memo_inbox),
            wx.AcceleratorEntry(wx.ACCEL_CTRL | wx.ACCEL_SHIFT, ord("M"), self.id_memo_compose),
            wx.AcceleratorEntry(wx.ACCEL_CTRL | wx.ACCEL_SHIFT, ord("E"), self.id_mail_compose),
            wx.AcceleratorEntry(wx.ACCEL_CTRL | wx.ACCEL_SHIFT, ord("N"), self.id_memo_check_now),
            wx.AcceleratorEntry(wx.ACCEL_ALT, ord("G"), self.id_goto),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F1, self.id_about),
            wx.AcceleratorEntry(wx.ACCEL_SHIFT, wx.WXK_F1, self.id_manual),
            wx.AcceleratorEntry(wx.ACCEL_ALT, ord("U"), self.id_update_check),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F5, self.id_board_refresh),
            wx.AcceleratorEntry(wx.ACCEL_NORMAL, wx.WXK_F7, self.id_settings),
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

    def on_show_settings(self, event):
        """F7: 통합 설정 대화상자 (테마 + 사운드 + 알림 + 업데이트)."""
        try:
            from settings_dialog import SettingsDialog
            dlg = SettingsDialog(self)
            result = dlg.ShowModal()
            dlg.Destroy()
            self._apply_full_theme()
            # 알림 주기 변경 반영
            if result == wx.ID_OK:
                try:
                    self.restart_memo_notifier()
                except Exception:
                    pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            speak(f"설정 대화상자를 여는 중 오류가 발생했습니다. {e}")
            wx.MessageBox(
                f"설정 대화상자 오류:\n{e}\n\n{traceback.format_exc()}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

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
                try:
                    from sound import play_event
                    play_event("page_move")
                except Exception:
                    pass
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
        try:
            from sound import play_event
            play_event("main_menu_return")
        except Exception:
            pass

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

        # Shift+F1: 사용자 설명서 / F1: 프로그램 정보
        elif keycode == wx.WXK_F1:
            if event.ShiftDown():
                self.on_show_manual(None)
            else:
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
                try:
                    from sound import play_event
                    play_event("home_end")
                except Exception:
                    pass

        # End: 마지막 항목으로 이동
        elif keycode == wx.WXK_END and not alt and not ctrl:
            if self.current_items:
                self._jump_to_line_silent(len(self.current_items) - 1)
                try:
                    from sound import play_event
                    play_event("home_end")
                except Exception:
                    pass

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

        type_names = [t[0] for t in search_types]
        type_label = wx.StaticText(panel, label="검색 유형(&T):")
        type_combo = wx.ComboBox(
            panel, choices=type_names,
            style=wx.CB_READONLY, name="검색 유형",
        )
        type_combo.SetSelection(0)

        hist_hint = " (↑↓: 최근 검색어)" if load_search_history() else ""
        query_label = wx.StaticText(panel, label=f"검색어(&S){hist_hint}:")
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
        dlg.SetMinSize(wx.Size(360, -1))
        dlg.Fit()
        query_input.SetFocus()
        dlg.Centre()

        # ── 검색 히스토리 (최대 10개, 최신이 0) ──
        history = load_search_history()
        # hist_idx: -1 = 사용자가 입력 중인 원본, 0.. = 히스토리 항목
        hist_state = {"idx": -1, "draft": ""}

        def apply_history(idx: int):
            """idx에 맞게 입력창/유형을 채운다."""
            hist_state["idx"] = idx
            if idx == -1:
                query_input.ChangeValue(hist_state["draft"])
            else:
                item = history[idx]
                query_input.ChangeValue(item.get("query", ""))
                t = item.get("type", "")
                if t in type_names:
                    type_combo.SetSelection(type_names.index(t))
                speak(f"최근 검색 {idx + 1}번, {item.get('query', '')}")
            query_input.SetInsertionPointEnd()

        def on_query_key(event: wx.KeyEvent):
            key = event.GetKeyCode()
            if key == wx.WXK_UP:
                if not history:
                    speak("검색 히스토리가 없습니다.")
                    return
                if hist_state["idx"] == -1:
                    hist_state["draft"] = query_input.GetValue()
                new_idx = min(hist_state["idx"] + 1, len(history) - 1)
                if new_idx == hist_state["idx"]:
                    speak("가장 오래된 검색어입니다.")
                    return
                apply_history(new_idx)
            elif key == wx.WXK_DOWN:
                if hist_state["idx"] == -1:
                    return
                new_idx = hist_state["idx"] - 1
                apply_history(new_idx)
            else:
                event.Skip()

        query_input.Bind(wx.EVT_KEY_DOWN, on_query_key)
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

        add_search_history(query, search_types[type_idx][0])

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

        # NAS 자료실 연결: 마운트 상태에 따라 탐색기로 열거나, 자격증명 입력
        if menu_item.type == "nas":
            self._activate_nas_menu()
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

    def _on_menu_nas_connect(self, event):
        """메뉴바 '도구 > 초록등대 자료실 연결' 핸들러."""
        try:
            self._activate_nas_menu()
        except Exception as e:
            import traceback
            traceback.print_exc()
            speak(f"NAS 연결 중 오류가 발생했습니다. {e}")
            wx.MessageBox(
                f"NAS 연결 중 예기치 않은 오류가 발생했습니다.\n\n{e}\n\n"
                f"{traceback.format_exc()}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

    def _activate_nas_menu(self):
        """초록등대 자료실 연결. rclone + WinFSP 기반."""
        try:
            from nas import (
                get_mounted_drive, open_in_explorer,
                prompt_and_mount, load_nas_credentials, mount,
                delete_nas_credentials, _is_auth_error, _is_winfsp_missing,
                WINFSP_DOWNLOAD_URL,
            )
        except Exception as e:
            speak(f"NAS 모듈을 불러올 수 없습니다. {e}")
            return

        # 1) 이미 마운트: 탐색기로 열기
        drive = get_mounted_drive()
        if drive:
            if open_in_explorer(drive):
                speak(f"{drive} 드라이브를 엽니다.")
            else:
                speak("드라이브를 열 수 없습니다.")
            return

        # 2) 저장된 자격증명이 있으면 자동 시도
        creds = load_nas_credentials()
        if creds:
            user, pw = creds
            speak("초록등대 자료실에 연결 중입니다.")
            ok, info = mount(user, pw)
            if ok:
                self._notify_nas_connected()
                open_in_explorer(info)
                return
            if _is_winfsp_missing(info):
                self._prompt_winfsp_install(WINFSP_DOWNLOAD_URL)
                return
            if _is_auth_error(info):
                delete_nas_credentials()
                if not self._ask_reenter_credentials_on_auth_error(info):
                    return
                # 3) 재입력 플로우로 폴스루
            else:
                wx.MessageBox(
                    f"저장된 자격증명으로 연결에 실패했습니다.\n\n{info}",
                    "NAS 연결 실패", wx.OK | wx.ICON_WARNING, self,
                )
                return

        # 3) 자격증명 입력 대화상자 → 마운트
        ok, info = prompt_and_mount(self, speak_func=None)
        if ok:
            from nas import open_in_explorer as _open
            self._notify_nas_connected()
            _open(info)
            return
        if _is_winfsp_missing(info):
            self._prompt_winfsp_install(WINFSP_DOWNLOAD_URL)
            return
        if _is_auth_error(info):
            if self._ask_reenter_credentials_on_auth_error(info):
                self._activate_nas_menu()
            return
        if info and info != "취소되었습니다.":
            wx.MessageBox(
                f"NAS 연결 실패.\n{info}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

    def _ask_reenter_credentials_on_auth_error(self, info: str) -> bool:
        """HTTP 401 발생 시 원인 안내 + 재입력 여부를 묻는 대화상자.
        YES 면 True, NO/닫기 면 False."""
        msg = (
            f"{info}\n\n"
            "서버가 아이디 또는 비밀번호를 인증하지 못했습니다.\n\n"
            "다음을 확인해 주세요:\n"
            "1) 아이디/비밀번호 오타 (공백 포함 여부)\n"
            "2) DSM 2단계 인증이 활성화된 경우 DSM의 '앱 비밀번호'를 발급받아 "
            "사용해야 합니다 (DSM 제어판 → 개인 → 계정 → 2단계 인증 → 앱 비밀번호)\n"
            "3) DSM 제어판 → 파일 서비스 → WebDAV 에서 WebDAV 서비스와 "
            "HTTPS 포트(5006)가 활성화되어 있는지\n"
            "4) DSM 제어판 → 사용자 → 해당 계정 → 편집 → 응용 프로그램 "
            "권한에서 WebDAV 허용되어 있는지\n\n"
            "자격증명을 다시 입력하시겠습니까?"
        )
        r = wx.MessageBox(
            msg, "NAS 연결 실패 — 인증 오류",
            wx.YES_NO | wx.ICON_WARNING, self,
        )
        return r == wx.YES

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
            # NAS 마운트도 함께 해제
            try:
                from nas import unmount as _nas_unmount
                _nas_unmount()
            except Exception:
                pass
            speak("로그아웃되었습니다.")
            self.Close()

    # ── 프로그램 정보 (F1) ──

    # ── 자동 업데이트 ──

    def _auto_update_check(self):
        """시작 시 백그라운드 업데이트 확인. 설정으로 끌 수 있음.

        사용자가 선택한 주기(실행 때마다/1주/2주/1달) 이내면 건너뜀.
        건너뛰기 선택한 버전은 알림 안 띄움.
        """
        settings = load_update_settings()
        if not settings.get("check_on_startup", True):
            return
        interval_hours = get_update_interval_hours(
            settings.get("check_interval", "weekly")
        )
        if interval_hours > 0 and self._is_recently_checked(
            settings.get("last_check_iso", ""), interval_hours
        ):
            return
        self._run_update_check(manual=False)

    @staticmethod
    def _is_recently_checked(last_iso: str, min_interval_hours: float) -> bool:
        """마지막 체크가 min_interval_hours 이내면 True."""
        if not last_iso:
            return False
        try:
            from datetime import datetime, timedelta
            last = datetime.fromisoformat(last_iso)
            return datetime.now() - last < timedelta(hours=min_interval_hours)
        except (TypeError, ValueError):
            return False

    def on_manual_update_check(self, event):
        """도움말 > 업데이트 확인 / Ctrl+Shift+U."""
        speak("업데이트를 확인하는 중입니다.")
        self._run_update_check(manual=True)

    def _run_update_check(self, manual: bool):
        """실제 조회는 백그라운드 스레드에서.

        manual=True 이면 "최신 버전입니다" / 네트워크 실패 메시지도 표시.
        manual=False 이면 새 버전 있을 때만 알림.
        """
        channel = load_update_settings().get("channel", "stable")

        def worker():
            info = check_latest_release(channel=channel)
            wx.CallAfter(self._on_update_check_done, info, manual)

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _on_update_check_done(self, info, manual: bool):
        if info is None:
            if manual:
                from updater import get_last_check_error
                reason = get_last_check_error() or "알 수 없는 오류"
                speak("업데이트 정보를 확인할 수 없습니다.")
                wx.MessageBox(
                    "업데이트 서버에 연결할 수 없습니다.\n"
                    "인터넷 연결을 확인한 뒤 다시 시도해 주세요.\n\n"
                    f"상세: {reason}",
                    "업데이트 확인 실패",
                    wx.OK | wx.ICON_WARNING, self,
                )
            return

        # 성공적인 조회는 last_check 에 기록 (하루 1회 제한 판단용)
        from datetime import datetime
        s = load_update_settings()
        s["last_check_iso"] = datetime.now().isoformat(timespec="seconds")
        save_update_settings(s)

        if not is_newer(APP_VERSION, info.version):
            if manual:
                speak(f"현재 버전이 최신입니다. {APP_VERSION}")
                wx.MessageBox(
                    f"현재 사용 중인 버전이 최신입니다.\n"
                    f"설치 버전: {APP_VERSION}\n"
                    f"최신 버전: {info.version}",
                    "최신 버전 사용 중",
                    wx.OK | wx.ICON_INFORMATION, self,
                )
            return

        # 자동 체크 + 이 버전 건너뛰기 선택한 경우는 조용히 종료
        if not manual:
            skipped = s.get("skip_version", "")
            if skipped and skipped == info.version:
                return

        # 릴리스 노트를 사용자 친화적으로 정리
        info.body = clean_release_notes(info.body)
        self._show_update_dialog(info)

    def _show_update_dialog(self, info):
        """새 버전 안내 대화상자."""
        dlg = UpdateDialog(self, info, APP_VERSION, self.current_font_size)
        try:
            apply_theme(dlg, make_font(self.current_font_size))
        except Exception:
            pass
        speak(f"새 버전 {info.version}이 있습니다. 업데이트하시겠습니까?")
        result = dlg.ShowModal()
        dlg.Destroy()

        if result == UpdateDialog.RESULT_UPDATE_NOW:
            self._start_update_download(info)
        elif result == UpdateDialog.RESULT_SKIP_VERSION:
            s = load_update_settings()
            s["skip_version"] = info.version
            save_update_settings(s)
            speak(f"버전 {info.version}을 건너뜁니다.")
        # RESULT_LATER 또는 취소는 아무것도 저장하지 않음

    def _start_update_download(self, info: ReleaseInfo):
        """설치 종류에 따라 설치 파일 / ZIP / 델타 중 적절한 경로 선택."""
        kind = detect_installation_kind()

        # 설치형: 기존 .exe 설치 경로 사용
        if kind == "installed" and info.installer_url:
            self._download_and_install_exe(info)
            return

        # 포터블: manifest 가 있으면 델타, 없으면 ZIP 전체 교체
        if kind == "portable":
            if info.manifest_url and info.zip_url:
                self._download_portable_delta(info)
                return
            if info.zip_url:
                self._download_portable_full(info)
                return

        # 설치형인데 installer_url 이 없거나, 포터블인데 zip 도 없을 때
        if info.installer_url:
            self._download_and_install_exe(info)
            return

        speak("다운로드 가능한 업데이트 파일이 없습니다.")
        url = info.html_url or UPDATE_RELEASES_PAGE
        wx.MessageBox(
            "이 릴리스에 자동 설치 가능한 파일이 첨부되지 않았습니다.\n"
            "브라우저로 릴리스 페이지를 열어 직접 받아 주세요.",
            "설치 파일 없음", wx.OK | wx.ICON_WARNING, self,
        )
        webbrowser.open(url)

    def _download_and_install_exe(self, info: ReleaseInfo):
        """설치형 업데이트: .exe 다운로드 후 SILENT 설치."""
        dest_dir = get_download_dir()
        file_name = info.installer_name or f"chorok_multi_{info.version}_setup.exe"
        dest_path = os.path.join(dest_dir, file_name)

        total_mb = info.installer_size / (1024 * 1024) if info.installer_size else 0
        initial_msg = (
            f"{file_name} 다운로드 준비 중..."
            if total_mb == 0
            else f"{file_name}\n약 {total_mb:.1f} MB 다운로드 준비 중..."
        )
        progress_dlg = wx.ProgressDialog(
            "업데이트 다운로드",
            initial_msg,
            maximum=1000,
            parent=self,
            style=wx.PD_APP_MODAL | wx.PD_CAN_ABORT
                  | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME
                  | wx.PD_AUTO_HIDE,
        )

        state = {"cancelled": False, "error": None, "path": None}

        def progress_cb(downloaded: int, total: int) -> bool:
            # 워커 스레드에서 호출됨. UI 갱신은 CallAfter.
            def ui_update():
                if state["cancelled"]:
                    return
                if total > 0:
                    ratio = downloaded / total
                    value = min(999, int(ratio * 1000))
                    msg = (
                        f"{file_name}\n"
                        f"{downloaded / (1024*1024):.1f} MB / "
                        f"{total / (1024*1024):.1f} MB"
                    )
                else:
                    value = 0
                    msg = f"{file_name}\n{downloaded / (1024*1024):.1f} MB 다운로드 중..."
                keep_going, _ = progress_dlg.Update(value, msg)
                if not keep_going:
                    state["cancelled"] = True
            wx.CallAfter(ui_update)
            return not state["cancelled"]

        def worker():
            try:
                path = download_installer(
                    info.installer_url, dest_path, progress_cb=progress_cb,
                )
                state["path"] = path
            except DownloadCancelled:
                state["cancelled"] = True
            except Exception as e:
                state["error"] = str(e)
            finally:
                wx.CallAfter(on_finished)

        def on_finished():
            try:
                progress_dlg.Destroy()
            except Exception:
                pass

            if state["cancelled"]:
                speak("업데이트를 취소했습니다.")
                return

            if state["error"]:
                speak("업데이트 다운로드에 실패했습니다.")
                wx.MessageBox(
                    f"설치 파일을 내려받지 못했습니다.\n{state['error']}\n\n"
                    f"브라우저에서 직접 내려받으시려면 '확인'을 누르세요.",
                    "다운로드 실패", wx.OK | wx.ICON_ERROR, self,
                )
                webbrowser.open(info.html_url or UPDATE_RELEASES_PAGE)
                return

            if state["path"] and os.path.exists(state["path"]):
                if not self._verify_download(state["path"], info):
                    return
                self._launch_installer(state["path"])

        speak("업데이트를 내려받는 중입니다.")
        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _verify_download(self, path: str, info: ReleaseInfo) -> bool:
        """다운로드 파일의 SHA256 을 릴리스 체크섬과 비교.

        - 체크섬 자산이 없으면 검증 건너뛰고 True 반환 (기존 호환성).
        - 검증 실패 시 파일 삭제, 사용자 경고, False 반환.
        """
        # 크기 1차 검증 (content-length 기준)
        if info.installer_size > 0:
            try:
                actual = os.path.getsize(path)
            except OSError:
                actual = 0
            if actual != info.installer_size:
                speak("다운로드된 파일의 크기가 일치하지 않습니다.")
                wx.MessageBox(
                    f"다운로드된 파일 크기가 예상과 다릅니다.\n"
                    f"예상: {info.installer_size:,} 바이트\n"
                    f"실제: {actual:,} 바이트\n\n"
                    f"다시 시도해 주세요.",
                    "다운로드 검증 실패", wx.OK | wx.ICON_ERROR, self,
                )
                try: os.remove(path)
                except OSError: pass
                return False

        # 체크섬 검증 (자산이 있을 때만)
        if not info.checksum_url:
            return True
        try:
            expected = fetch_expected_checksum(info.checksum_url, info.installer_name)
        except Exception:
            expected = None
        if not expected:
            # 체크섬 파일은 있었지만 파싱 실패 — 조용히 스킵
            return True
        try:
            actual_hash = sha256_of_file(path)
        except OSError as e:
            wx.MessageBox(
                f"설치 파일을 읽는 중 오류가 발생했습니다.\n{e}",
                "검증 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return False
        if actual_hash.lower() != expected.lower():
            speak("다운로드된 파일의 무결성 검증에 실패했습니다.")
            wx.MessageBox(
                "다운로드된 파일의 SHA256 체크섬이 릴리스와 일치하지 않습니다.\n"
                "파일이 전송 중 손상되었거나 변조되었을 수 있습니다.\n"
                "파일을 삭제합니다. 다시 시도해 주세요.\n\n"
                f"예상: {expected}\n"
                f"실제: {actual_hash}",
                "무결성 검증 실패", wx.OK | wx.ICON_ERROR, self,
            )
            try: os.remove(path)
            except OSError: pass
            return False
        return True

    def _launch_installer(self, installer_path: str):
        """내려받은 설치 파일 실행. 사용자 확인 후 초록멀티를 종료."""
        ans = wx.MessageBox(
            "다운로드가 완료되었습니다.\n"
            "설치 프로그램을 실행하면 초록멀티가 종료되고\n"
            "새 버전이 설치됩니다. 지금 설치하시겠습니까?",
            "업데이트 설치",
            wx.YES_NO | wx.ICON_INFORMATION, self,
        )
        if ans != wx.YES:
            speak("설치 파일이 준비되었습니다. 다음 실행 시 다시 안내할 수 있습니다.")
            return

        try:
            # Inno Setup 인자: /SILENT(마법사 없이 설치 진행), /CLOSEAPPLICATIONS
            # (실행 중인 초록멀티를 설치 프로그램이 자동 종료), /RESTARTAPPLICATIONS
            # (설치 후 자동 재시작). 비-Inno 설치 파일이어도 해당 인자는 무시됨.
            import subprocess
            subprocess.Popen(
                [installer_path, "/SILENT", "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS"],
                close_fds=True,
            )
        except Exception as e:
            wx.MessageBox(
                f"설치 프로그램을 실행하지 못했습니다.\n{e}\n\n"
                f"내려받은 파일 위치:\n{installer_path}",
                "설치 실행 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return

        speak("설치를 시작합니다. 초록멀티를 종료합니다.")
        # 짧게 대기 후 프레임 닫기 → OnExit → MainLoop 종료
        wx.CallLater(800, self.Close)

    # ── 포터블 업데이트 ──

    def _download_portable_full(self, info: ReleaseInfo):
        """포터블 전체 ZIP 다운로드 → 추출 → 재시작 스크립트."""
        dest_dir = get_download_dir()
        file_name = info.zip_name or f"chorok_multi_{info.version}.zip"
        dest_path = os.path.join(dest_dir, file_name)

        total_mb = info.zip_size / (1024 * 1024) if info.zip_size else 0
        initial_msg = (
            f"{file_name} 다운로드 준비 중..."
            if total_mb == 0
            else f"{file_name}\n약 {total_mb:.1f} MB 다운로드 준비 중..."
        )
        progress_dlg = wx.ProgressDialog(
            "포터블 업데이트 다운로드",
            initial_msg,
            maximum=1000, parent=self,
            style=wx.PD_APP_MODAL | wx.PD_CAN_ABORT
                  | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME
                  | wx.PD_AUTO_HIDE,
        )
        state = {"cancelled": False, "error": None, "path": None}

        def progress_cb(downloaded, total):
            def ui():
                if state["cancelled"]:
                    return
                if total > 0:
                    value = min(999, int(downloaded / total * 1000))
                    msg = (
                        f"{file_name}\n"
                        f"{downloaded / (1024*1024):.1f} MB / "
                        f"{total / (1024*1024):.1f} MB"
                    )
                else:
                    value = 0
                    msg = f"{file_name}\n{downloaded / (1024*1024):.1f} MB 다운로드 중..."
                keep_going, _ = progress_dlg.Update(value, msg)
                if not keep_going:
                    state["cancelled"] = True
            wx.CallAfter(ui)
            return not state["cancelled"]

        def worker():
            try:
                path = download_installer(info.zip_url, dest_path, progress_cb=progress_cb)
                state["path"] = path
            except DownloadCancelled:
                state["cancelled"] = True
            except Exception as e:
                state["error"] = str(e)
            finally:
                wx.CallAfter(on_finished)

        def on_finished():
            try:
                progress_dlg.Destroy()
            except Exception:
                pass
            if state["cancelled"]:
                speak("업데이트를 취소했습니다.")
                return
            if state["error"]:
                speak("포터블 업데이트 다운로드에 실패했습니다.")
                wx.MessageBox(
                    f"ZIP 파일을 내려받지 못했습니다.\n{state['error']}",
                    "다운로드 실패", wx.OK | wx.ICON_ERROR, self,
                )
                return
            self._apply_portable_update(info, state["path"], delta_paths=None)

        speak("포터블 업데이트를 내려받는 중입니다.")
        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _download_portable_delta(self, info: ReleaseInfo):
        """델타 업데이트: manifest 비교 → 변경된 파일만 ZIP 에서 추출."""
        speak("변경된 파일 목록을 확인 중입니다.")

        def worker():
            manifest = fetch_manifest(info.manifest_url)
            if manifest is None:
                wx.CallAfter(
                    self._on_delta_fallback, info,
                    "매니페스트를 불러올 수 없어 전체 ZIP 업데이트로 전환합니다.",
                )
                return
            install_dir = get_install_dir()
            delta = compute_delta(install_dir, manifest)
            wx.CallAfter(self._on_delta_ready, info, manifest, delta)

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _on_delta_fallback(self, info: ReleaseInfo, reason: str):
        speak(reason)
        self._download_portable_full(info)

    def _on_delta_ready(self, info: ReleaseInfo, manifest: dict, delta: list):
        if not delta:
            speak("이미 최신 상태입니다.")
            wx.MessageBox(
                "변경된 파일이 없습니다. 이미 최신 상태입니다.",
                "업데이트 불필요", wx.OK | wx.ICON_INFORMATION, self,
            )
            # last_check 도 이미 갱신되었으므로 종료
            return

        total_files = len(manifest.get("files") or [])
        delta_files = len(delta)
        # 델타가 전체의 70% 넘으면 전체 ZIP 교체가 효율적
        if total_files > 0 and delta_files / total_files > 0.7:
            speak("변경된 파일이 많아 전체 업데이트로 전환합니다.")
            self._download_portable_full(info)
            return

        ans = wx.MessageBox(
            f"변경된 파일 {delta_files}개를 내려받아 교체합니다.\n"
            f"(전체 {total_files}개 중 {delta_files}개)\n\n"
            f"계속하시겠습니까?",
            "델타 업데이트",
            wx.YES_NO | wx.ICON_INFORMATION, self,
        )
        if ans != wx.YES:
            return

        delta_paths = [d["path"] for d in delta]
        self._download_portable_full_inner(info, delta_paths=delta_paths, manifest=manifest)

    def _apply_portable_update(
        self, info: ReleaseInfo, zip_path, delta_paths=None, manifest=None,
    ):
        """내려받은 ZIP 을 풀고 재시작 스크립트 실행.

        delta_paths 가 None 이면 전체 추출, 리스트면 해당 파일만 추출.
        """
        if not zip_path or not os.path.exists(zip_path):
            wx.MessageBox(
                "내려받은 파일을 찾을 수 없습니다.", "업데이트 실패",
                wx.OK | wx.ICON_ERROR, self,
            )
            return

        install_dir = get_install_dir()
        staging_dir = os.path.join(get_download_dir(), f"staging_{info.version}")
        # 이전 staging 정리
        try:
            if os.path.isdir(staging_dir):
                import shutil
                shutil.rmtree(staging_dir, ignore_errors=True)
            os.makedirs(staging_dir, exist_ok=True)
        except OSError as e:
            wx.MessageBox(
                f"임시 폴더를 준비하지 못했습니다.\n{e}",
                "업데이트 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return

        # 압축 해제 진행률 창
        extract_dlg = wx.ProgressDialog(
            "업데이트 압축 해제",
            "파일을 풀어내는 중입니다...",
            maximum=1000, parent=self,
            style=wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME | wx.PD_AUTO_HIDE,
        )
        extract_state = {"cancelled": False, "error": None}

        def xprog(done, total):
            def ui():
                if total > 0:
                    extract_dlg.Update(
                        min(999, int(done / total * 1000)),
                        f"{done} / {total} 파일 처리 중",
                    )
                else:
                    extract_dlg.Pulse(f"{done} 파일 처리 중")
            wx.CallAfter(ui)
            return not extract_state["cancelled"]

        def work():
            try:
                extract_zip(zip_path, staging_dir, only_paths=delta_paths, progress_cb=xprog)
            except Exception as e:
                extract_state["error"] = str(e)
            finally:
                wx.CallAfter(done)

        def done():
            try:
                extract_dlg.Destroy()
            except Exception:
                pass
            if extract_state["error"]:
                speak("압축 해제에 실패했습니다.")
                wx.MessageBox(
                    f"압축 해제에 실패했습니다.\n{extract_state['error']}",
                    "업데이트 실패", wx.OK | wx.ICON_ERROR, self,
                )
                return
            self._finalize_portable_update(info, staging_dir, manifest=manifest)

        import threading
        threading.Thread(target=work, daemon=True).start()

    def _download_portable_full_inner(self, info: ReleaseInfo, delta_paths=None, manifest=None):
        """ZIP 다운로드 + 델타 추출 (_download_portable_full 과 달리
        추출 대상 경로를 넘길 수 있음)."""
        dest_dir = get_download_dir()
        file_name = info.zip_name or f"chorok_multi_{info.version}.zip"
        dest_path = os.path.join(dest_dir, file_name)

        total_mb = info.zip_size / (1024 * 1024) if info.zip_size else 0
        initial_msg = (
            f"{file_name} 다운로드 준비 중..."
            if total_mb == 0
            else f"{file_name}\n약 {total_mb:.1f} MB 다운로드 준비 중..."
        )
        progress_dlg = wx.ProgressDialog(
            "업데이트 다운로드",
            initial_msg,
            maximum=1000, parent=self,
            style=wx.PD_APP_MODAL | wx.PD_CAN_ABORT
                  | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME
                  | wx.PD_AUTO_HIDE,
        )
        state = {"cancelled": False, "error": None, "path": None}

        def progress_cb(downloaded, total):
            def ui():
                if state["cancelled"]:
                    return
                if total > 0:
                    value = min(999, int(downloaded / total * 1000))
                    msg = (
                        f"{file_name}\n"
                        f"{downloaded / (1024*1024):.1f} MB / "
                        f"{total / (1024*1024):.1f} MB"
                    )
                else:
                    value = 0
                    msg = f"{file_name}\n{downloaded / (1024*1024):.1f} MB 다운로드 중..."
                keep_going, _ = progress_dlg.Update(value, msg)
                if not keep_going:
                    state["cancelled"] = True
            wx.CallAfter(ui)
            return not state["cancelled"]

        def worker():
            try:
                path = download_installer(info.zip_url, dest_path, progress_cb=progress_cb)
                state["path"] = path
            except DownloadCancelled:
                state["cancelled"] = True
            except Exception as e:
                state["error"] = str(e)
            finally:
                wx.CallAfter(on_finished)

        def on_finished():
            try:
                progress_dlg.Destroy()
            except Exception:
                pass
            if state["cancelled"]:
                speak("업데이트를 취소했습니다.")
                return
            if state["error"]:
                speak("업데이트 다운로드에 실패했습니다.")
                wx.MessageBox(
                    f"ZIP 파일을 내려받지 못했습니다.\n{state['error']}",
                    "다운로드 실패", wx.OK | wx.ICON_ERROR, self,
                )
                return
            self._apply_portable_update(info, state["path"],
                                        delta_paths=delta_paths, manifest=manifest)

        speak("변경된 파일을 내려받는 중입니다." if delta_paths else "업데이트를 내려받는 중입니다.")
        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _finalize_portable_update(self, info: ReleaseInfo, staging_dir: str, manifest=None):
        """재시작 스크립트 실행 후 초록멀티 종료."""
        # manifest 가 있으면 새 exe 이름을 거기서, 없으면 기존 이름 유지
        if manifest and manifest.get("executable"):
            new_exe_name = str(manifest.get("executable"))
        else:
            new_exe_name = os.path.basename(sys.executable) if getattr(sys, "frozen", False) else "초록멀티 v1.4.exe"
        old_exe_path = sys.executable if getattr(sys, "frozen", False) else os.path.join(get_install_dir(), new_exe_name)

        ans = wx.MessageBox(
            "다운로드가 완료되었습니다.\n"
            "초록멀티를 종료하고 파일 교체 후 새 버전을 실행합니다.\n"
            "계속하시겠습니까?",
            "포터블 업데이트 적용",
            wx.YES_NO | wx.ICON_INFORMATION, self,
        )
        if ans != wx.YES:
            speak("업데이트 적용을 취소했습니다. 스테이징 폴더에 새 파일이 남아 있습니다.")
            return

        # staging_dir 검증 — 새 exe 가 실제로 들어있는지 확인
        expected_new_exe = os.path.join(staging_dir, new_exe_name)
        if not os.path.exists(expected_new_exe):
            # 최상위 폴더가 남아있을 수 있음 (extract_zip 의 strip_prefix 가 작동 안 한 경우)
            alt = None
            try:
                for entry in os.listdir(staging_dir):
                    candidate = os.path.join(staging_dir, entry, new_exe_name)
                    if os.path.exists(candidate):
                        alt = os.path.join(staging_dir, entry)
                        break
            except OSError:
                pass
            if alt is None:
                wx.MessageBox(
                    f"업데이트 파일이 손상되었습니다.\n"
                    f"새 실행 파일이 스테이징 폴더에 없습니다.\n\n"
                    f"새 exe 이름: {new_exe_name}\n"
                    f"스테이징 폴더:\n{staging_dir}\n\n"
                    f"스테이징 폴더를 열어 내용을 확인해 주세요.",
                    "업데이트 파일 손상", wx.OK | wx.ICON_ERROR, self,
                )
                return
            # 최상위 폴더가 남아있는 경우 robocopy 소스를 그 폴더로 조정
            staging_dir = alt

        # 실행 중 EXE 를 .old 로 rename 시도. 성공하면 경쟁 조건 없이 교체 가능.
        backup_exe_path = self._rename_running_exe_to_backup(old_exe_path)

        try:
            script = write_restart_script(
                staging_dir=staging_dir,
                install_dir=get_install_dir(),
                new_exe_name=new_exe_name,
                old_exe_path=old_exe_path,
                backup_exe_path=backup_exe_path,
            )
            import subprocess
            # CREATE_NO_WINDOW = 0x08000000 — 콘솔은 숨기되 PS 에게 콘솔 자체는 주어야
            # 내부의 robocopy/Start-Process 호출이 정상 동작한다. DETACHED_PROCESS(0x8)
            # 는 PS 에 콘솔이 아예 없어서 -NoNewWindow 자식이 안 돌아간다.
            subprocess.Popen(
                [
                    "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                    "-WindowStyle", "Hidden",
                    "-File", script,
                ],
                creationflags=0x08000000,  # CREATE_NO_WINDOW
                close_fds=True,
            )
        except Exception as e:
            wx.MessageBox(
                f"업데이트 스크립트를 실행하지 못했습니다.\n{e}\n\n"
                f"스테이징 폴더:\n{staging_dir}",
                "업데이트 실행 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return

        speak("업데이트를 적용하고 초록멀티를 재시작합니다.")
        # rename 이 선행됐으면 구 exe 경쟁이 없으므로 빠르게 Close.
        # rename 실패 시에도 PS 쪽 Start-Sleep 이 Close 완료를 커버.
        wx.CallLater(200, self.Close)

    def _rename_running_exe_to_backup(self, old_exe_path: str) -> "str | None":
        """실행 중인 EXE 를 .old 로 rename. 성공 시 .old 경로, 실패 시 None.

        Windows NTFS 는 in-use 파일의 rename 을 허용하지만, 드물게 백신·
        스크린리더가 추가 핸들을 잡고 있으면 SHARING_VIOLATION 이 발생.
        실패 시 None 을 반환해서 호출자가 기존 삭제-재시도 경로로 폴백.
        """
        if not getattr(sys, "frozen", False):
            return None
        backup_path = old_exe_path + ".old"
        try:
            import ctypes
            from ctypes import wintypes
            MoveFileExW = ctypes.windll.kernel32.MoveFileExW
            MoveFileExW.argtypes = (wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD)
            MoveFileExW.restype = wintypes.BOOL
            MOVEFILE_REPLACE_EXISTING = 0x1
            ok = MoveFileExW(old_exe_path, backup_path, MOVEFILE_REPLACE_EXISTING)
            if not ok:
                return None
            return backup_path
        except Exception:
            return None

    # ── 사용자 설명서 (Shift+F1) ──

    def on_show_manual(self, event):
        """Shift+F1: 사용자 설명서 대화상자"""
        manual_path = resource_path("data", "manual.txt")
        try:
            with open(manual_path, "r", encoding="utf-8") as f:
                text = f.read()
        except OSError as e:
            speak("사용자 설명서 파일을 열 수 없습니다.")
            wx.MessageBox(
                f"사용자 설명서를 불러올 수 없습니다.\n"
                f"경로: {manual_path}\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )
            return

        chapters = _parse_manual_chapters(text)
        if not chapters:
            speak("사용자 설명서 내용을 읽을 수 없습니다.")
            return

        dlg = ManualDialog(self, chapters, self.current_font_size)
        try:
            apply_theme(dlg, make_font(self.current_font_size))
        except Exception:
            pass
        speak("사용자 설명서. 왼쪽 목록에서 챕터를 선택하세요.")
        dlg.ShowModal()
        dlg.Destroy()

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
            "Ctrl+N: 초록등대 자료실(NAS) 연결",
            "Ctrl+M: 쪽지함 열기",
            "Ctrl+Shift+M: 쪽지 쓰기",
            "Ctrl+Shift+E: 메일함 열기",
            "Ctrl+Shift+N: 알림 센터 열기 (새 쪽지·메일)",
            "Alt+E: 관리자에게 메일 보내기",
            "F1: 프로그램 정보",
            "Shift+F1: 사용자 설명서",
            "Alt+U: 업데이트 확인",
            "Alt+F4: 프로그램 종료",
            "",
            "=== 쪽지함·메일함 내부 ===",
            "↑/↓: 항목 이동",
            "Enter: 쪽지/메일 열기",
            "D 또는 Delete: 선택 항목 삭제",
            "Shift+Delete: 현재 함 전체 비우기",
            "R: 답장 (받은함)",
            "N: 새 쪽지/메일 작성",
            "F: 새로고침",
            "PageDown/PageUp: 다음 페이지 누적 로드",
            "Alt+R / Alt+S: 받은함 / 보낸함 전환",
            "Alt+A: 모든 쪽지/메일 삭제",
            "",
            "=== 쪽지·메일 보기 창 ===",
            "PageUp/PageDown, Alt+P/Alt+N: 이전/다음 항목",
            "R: 답장  D/Delete: 삭제  Esc: 닫기",
            "(메일) B: 본문 저장  Alt+S: 첨부 선택 저장  Alt+Shift+S: 모든 첨부 저장",
            "",
            "=== 알림 센터 (Ctrl+Shift+N) ===",
            "Enter: 선택 항목 열기",
            "D 또는 Delete: 선택 알림 지우기",
            "A: 모든 알림 지우기",
            "F: 새로고침",
            "",
            "=== 화면 설정 (저시력 지원) ===",
            "F7: 설정 창 열기 (테마·글꼴·사운드 통합)",
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
        """도움말 > 관리자에게 메일 보내기 (Alt+E) — 소리샘 세션으로 formmail 발송."""
        try:
            from mail import MailWriteDialog
            dlg = MailWriteDialog(self, self.session, mode=MailWriteDialog.MODE_ADMIN)
            dlg.ShowModal()
            dlg.Destroy()
        except Exception as e:
            speak("메일 대화상자 호출 중 오류가 발생했습니다.")
            wx.MessageBox(f"메일 대화상자 호출 중 오류가 발생했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    def on_open_memo_inbox(self, event):
        """도구 > 쪽지함 열기 (Ctrl+M)."""
        try:
            from memo import MemoInboxDialog
            dlg = MemoInboxDialog(self, self.session)
            dlg.ShowModal()
            dlg.Destroy()
            # 쪽지함을 직접 열었으니 현재 안 읽은 쪽지는 모두 "본 것" 으로 처리
            self._unread_memo_count = 0
            self._update_title_unread()
            if hasattr(self, "_memo_notifier") and self._memo_notifier:
                self._memo_notifier.mark_all_as_seen()
        except Exception as e:
            speak("쪽지함을 여는 중 오류가 발생했습니다.")
            wx.MessageBox(f"쪽지함을 여는 중 오류가 발생했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    def on_memo_check_now(self, event):
        """도구 > 알림 센터 열기 (Ctrl+Shift+N).

        열기 전에 쪽지·메일 최신 상태를 한 번 동기식으로 당겨서 센터에 채운 뒤
        대화상자 표시. 현재 안 읽은 항목은 모두 알림 센터에 들어감.
        """
        speak("알림 센터를 엽니다.")
        self._populate_notification_center_sync()
        from notification_dialog import NotificationCenterDialog
        dlg = NotificationCenterDialog(
            self, self.session,
            on_open_memo=self._open_memo_from_notification,
            on_open_mail=self._open_mail_from_notification,
        )
        dlg.ShowModal()
        dlg.Destroy()

    def _populate_notification_center_sync(self):
        """현재 안 읽은 쪽지·메일을 알림 센터에 즉시 반영 (동기, UI 스레드)."""
        from notification import NotificationItem, get_center
        center = get_center()
        # 쪽지
        try:
            from memo import fetch_inbox
            ok, items = fetch_inbox(self.session, kind="recv")
            if ok and isinstance(items, list):
                to_add = []
                for it in items:
                    if it.is_read:
                        continue
                    to_add.append(NotificationItem(
                        type="memo",
                        item_id=it.me_id,
                        sender=it.counterpart,
                        summary=it.summary,
                        timestamp=it.date,
                        extra=it,
                    ))
                center.add_many(to_add)
        except Exception:
            pass
        # 메일 — 받은함만 (새 메일 알림은 받은함 기준)
        try:
            from mail import fetch_mail_list
            ok, mitems = fetch_mail_list(self.session, kind="recv")
            if ok and isinstance(mitems, list):
                to_add = []
                for it in mitems:
                    if it.is_read:
                        continue
                    to_add.append(NotificationItem(
                        type="mail",
                        item_id=it.mail_id,
                        sender=it.sender,
                        summary=it.subject,
                        timestamp=it.date,
                        extra=it,
                    ))
                center.add_many(to_add)
        except Exception:
            pass

    def _open_memo_from_notification(self, notif):
        """알림 센터에서 쪽지 알림을 열 때 호출."""
        try:
            from memo import fetch_memo, MemoViewDialog
            ok, content = fetch_memo(self.session, notif.item_id, kind="recv")
            if not ok:
                speak("쪽지를 불러오지 못했습니다.")
                wx.MessageBox(f"쪽지를 불러오지 못했습니다.\n{content}",
                              "오류", wx.OK | wx.ICON_ERROR, self)
                return
            dlg = MemoViewDialog(self, self.session, content)
            dlg.ShowModal()
            dlg.Destroy()
        except Exception as e:
            wx.MessageBox(f"쪽지 열기 실패.\n{e}", "오류",
                          wx.OK | wx.ICON_ERROR, self)

    def _open_mail_from_notification(self, notif):
        """알림 센터에서 메일 알림을 선택했을 때 — MailViewDialog 로 본문 표시."""
        try:
            from mail import fetch_mail_content, MailViewDialog
            ok, content = fetch_mail_content(self.session, notif.item_id, kind="recv")
            if not ok:
                speak("메일을 불러오지 못했습니다.")
                wx.MessageBox(f"메일을 불러오지 못했습니다.\n{content}",
                              "오류", wx.OK | wx.ICON_ERROR, self)
                return
            dlg = MailViewDialog(self, self.session, content)
            dlg.ShowModal()
            dlg.Destroy()
        except Exception as e:
            wx.MessageBox(f"메일 열기 실패.\n{e}", "오류",
                          wx.OK | wx.ICON_ERROR, self)

    # ── 쪽지 실시간 알림 ──

    def _start_memo_notifier(self):
        """백그라운드 알림 폴링 시작 — 쪽지(Memo) + 메일(Mail) 둘 다.

        사용자 설정에 따라 각각 on/off 가능. 주기는 공통.
        메일은 MemoNotifier 가 tick 마다 같이 poll_once_async 를 호출하는 방식.
        """
        if not self.session:
            return
        try:
            from memo import MemoNotifier
            from mail import MailNotifier
            from settings_dialog import load_notify_settings
            settings = load_notify_settings()
            interval = int(settings.get("interval_sec", 60))
            check_memo = bool(settings.get("check_memo", True))
            check_mail = bool(settings.get("check_mail", True))
            if interval <= 0 or (not check_memo and not check_mail):
                self._memo_notifier = None
                self._mail_notifier = None
                return
            self._mail_notifier = None
            if check_mail:
                self._mail_notifier = MailNotifier(self, self.session)
                self._mail_notifier.start_initial_fill()
            if check_memo:
                self._memo_notifier = MemoNotifier(self, self.session, self._on_new_memo_or_mail)
                # tick 시 mail 도 함께 체크하도록 hook
                self._memo_notifier._piggyback_mail = self._poll_mail_from_memo_tick
                # MemoNotifier 의 _check_in_bg 마지막에 piggyback 호출 필요 — wrap 대신
                # 여기서는 wx.Timer 로 별도 타이머 안 쓰고 MemoNotifier tick 이 돌면서
                # 메일도 같이 체크되게끔 wx.Timer 에 추가 Bind.
                import wx as _wx
                self._mail_timer = _wx.Timer(self)
                self.Bind(_wx.EVT_TIMER, self._on_mail_tick, self._mail_timer)
                self._mail_timer.Start(interval * 1000)
                self._memo_notifier.start(interval_sec=interval)
            else:
                # 쪽지 체크 안 하고 메일만
                import wx as _wx
                self._memo_notifier = None
                self._mail_timer = _wx.Timer(self)
                self.Bind(_wx.EVT_TIMER, self._on_mail_tick, self._mail_timer)
                self._mail_timer.Start(interval * 1000)
        except Exception:
            self._memo_notifier = None
            self._mail_notifier = None

    def _on_mail_tick(self, event):
        """메일 폴링 tick — 새 메일이 있으면 알림 센터에 추가 + 알림."""
        if not getattr(self, "_mail_notifier", None):
            return
        self._mail_notifier.poll_once_async(on_new_items=self._on_new_mail)

    def _poll_mail_from_memo_tick(self):
        """호환용 no-op."""
        pass

    def _on_new_memo_or_mail(self, count, new_items):
        """기존 _on_new_memo 의 래퍼 — 이름만 의미 명확히."""
        self._on_new_memo(count, new_items)

    def _on_new_mail(self, new_items):
        """새 메일 도착 — 알림 센터에 등록 + 사운드/TTS/대화상자."""
        count = len(new_items)
        if count == 0:
            return
        try:
            from notification import NotificationItem, get_center
            center = get_center()
            to_add = [
                NotificationItem(
                    type="mail", item_id=it.mail_id,
                    sender=it.sender, summary=it.subject,
                    timestamp=it.date, extra=it,
                ) for it in new_items
            ]
            center.add_many(to_add)
        except Exception:
            pass
        try:
            from sound import play_event
            play_event("memo_new")
        except Exception:
            pass
        self._unread_mail_count += count
        self._update_title_unread()
        sender = new_items[0].sender if new_items else "알 수 없음"
        if count == 1:
            speak(f"새 메일이 도착했습니다. 보낸 사람 {sender}")
        else:
            speak(f"새 메일이 {count}개 도착했습니다.")
        if count == 1:
            msg = (
                f"새 메일이 도착했습니다.\n"
                f"보낸 사람: {sender}\n"
                f"제목: {new_items[0].subject}\n\n"
                f"알림 센터를 여시겠습니까?"
            )
        else:
            msg = f"새 메일이 {count}개 도착했습니다.\n\n알림 센터를 여시겠습니까?"
        ans = wx.MessageBox(msg, "새 메일 도착",
                            wx.YES_NO | wx.ICON_INFORMATION, self)
        if ans == wx.YES:
            self.on_memo_check_now(None)

    def restart_memo_notifier(self):
        """설정 변경 후 호출 — 기존 타이머 중단 후 새 주기로 재시작."""
        try:
            if getattr(self, "_memo_notifier", None):
                self._memo_notifier.stop()
        except Exception:
            pass
        try:
            if getattr(self, "_mail_timer", None):
                self._mail_timer.Stop()
        except Exception:
            pass
        self._memo_notifier = None
        self._mail_notifier = None
        self._mail_timer = None
        self._start_memo_notifier()

    def _on_new_memo(self, count: int, new_items: list):
        """새 쪽지 도착 콜백 — 알림 센터에 등록 + 사운드·TTS·제목바 업데이트.

        자동 폴링에서 호출됨. 사용자가 원본을 열 때는 알림 센터에서 선택하여 연다.
        """
        # 1. 알림 센터에 등록
        try:
            from notification import NotificationItem, get_center
            center = get_center()
            to_add = [
                NotificationItem(
                    type="memo", item_id=it.me_id,
                    sender=it.counterpart, summary=it.summary,
                    timestamp=it.date, extra=it,
                ) for it in new_items
            ]
            center.add_many(to_add)
        except Exception:
            pass

        # 2. 사운드
        try:
            from sound import play_event
            play_event("memo_new")
        except Exception:
            pass

        # 3. 제목바 갱신
        self._unread_memo_count += count
        self._update_title_unread()

        # 4. TTS
        sender = new_items[0].counterpart if new_items else "알 수 없음"
        if count == 1:
            speak(f"새 쪽지가 도착했습니다. 보낸 사람 {sender}")
        else:
            speak(f"새 쪽지가 {count}개 도착했습니다.")

        # 5. 확인 대화상자 — Yes 면 알림 센터 오픈
        if count == 1:
            msg = (
                f"새 쪽지가 도착했습니다.\n"
                f"보낸 사람: {sender}\n\n"
                f"알림 센터를 여시겠습니까?"
            )
        else:
            msg = (
                f"새 쪽지가 {count}개 도착했습니다.\n\n"
                f"알림 센터를 여시겠습니까?"
            )
        ans = wx.MessageBox(msg, "새 쪽지 도착",
                            wx.YES_NO | wx.ICON_INFORMATION, self)
        if ans == wx.YES:
            self.on_memo_check_now(None)

    def _update_title_unread(self):
        """제목바에 안 읽은 쪽지·메일 개수 표시 (있을 때만)."""
        base = getattr(self, "_base_title", APP_NAME)
        memo_count = getattr(self, "_unread_memo_count", 0)
        mail_count = getattr(self, "_unread_mail_count", 0)
        parts = []
        if memo_count > 0:
            parts.append(f"새 쪽지 {memo_count}")
        if mail_count > 0:
            parts.append(f"새 메일 {mail_count}")
        if parts:
            self.SetTitle(f"{base} - {' · '.join(parts)}")
        else:
            self.SetTitle(base)

    def on_open_memo_compose(self, event):
        """도구 > 쪽지 쓰기 (Ctrl+Shift+M)."""
        try:
            from memo import MemoWriteDialog
            dlg = MemoWriteDialog(self, self.session)
            dlg.ShowModal()
            dlg.Destroy()
        except Exception as e:
            speak("쪽지 작성 중 오류가 발생했습니다.")
            wx.MessageBox(f"쪽지 작성 중 오류가 발생했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    def on_open_mail_compose(self, event):
        """도구 > 메일함 열기 (Ctrl+Shift+E) — 사이트 내 메일함.
        기존 "메일 보내기 (formmail)" 는 Alt+E 관리자 메일 메뉴에서 계속 사용 가능."""
        try:
            from mail import MailInboxDialog
            dlg = MailInboxDialog(self, self.session)
            dlg.ShowModal()
            dlg.Destroy()
            # 메일함을 직접 열었으니 안 읽은 메일 카운트 초기화
            self._unread_mail_count = 0
            self._update_title_unread()
            notifier = getattr(self, "_mail_notifier", None)
            if notifier is not None:
                try:
                    from mail import fetch_mail_list
                    ok, items = fetch_mail_list(self.session, kind="recv")
                    if ok and isinstance(items, list):
                        for it in items:
                            notifier.seen_ids.add(it.mail_id)
                except Exception:
                    pass
        except Exception as e:
            speak("메일함을 여는 중 오류가 발생했습니다.")
            wx.MessageBox(f"메일함을 여는 중 오류가 발생했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    def on_exit(self, event):
        self.Close()


# ─────────────────────────────────────────────────────────────
# 업데이트 안내 대화상자
# ─────────────────────────────────────────────────────────────

class UpdateDialog(wx.Dialog):
    """새 버전 발견 시 안내 대화상자.

    버튼 결과:
        RESULT_UPDATE_NOW   — 지금 업데이트
        RESULT_LATER        — 나중에 (아무것도 저장 안 함)
        RESULT_SKIP_VERSION — 이 버전 건너뛰기 (skip_version 저장)
    """
    RESULT_UPDATE_NOW = 1001
    RESULT_LATER = 1002
    RESULT_SKIP_VERSION = 1003

    def __init__(self, parent, info, current_version: str, font_size: int):
        super().__init__(
            parent,
            title="초록멀티 업데이트 알림",
            size=(560, 460),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self._info = info

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        heading = wx.StaticText(
            panel,
            label=f"새 버전 {info.version}이(가) 공개되었습니다.",
        )
        heading.SetFont(make_font(font_size + 2).Bold())

        sub = wx.StaticText(
            panel,
            label=f"현재 버전: {current_version}     최신 버전: {info.version}",
        )

        notes_label = wx.StaticText(panel, label="변경 사항(&N):")
        notes = wx.TextCtrl(
            panel,
            value=info.body or "(변경 사항이 비어 있습니다)",
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_DONTWRAP,
            name="변경 사항",
        )
        notes.SetFont(make_font(font_size))

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        update_btn = wx.Button(panel, label="지금 업데이트(&U)")
        later_btn = wx.Button(panel, label="나중에(&L)")
        skip_btn = wx.Button(panel, label="이 버전 건너뛰기(&S)")
        btn_sizer.Add(update_btn, 0, wx.ALL, 5)
        btn_sizer.Add(later_btn, 0, wx.ALL, 5)
        btn_sizer.Add(skip_btn, 0, wx.ALL, 5)

        sizer.Add(heading, 0, wx.ALL, 10)
        sizer.Add(sub, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        sizer.Add(notes_label, 0, wx.LEFT | wx.TOP, 10)
        sizer.Add(notes, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(sizer)

        update_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(self.RESULT_UPDATE_NOW))
        later_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(self.RESULT_LATER))
        skip_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(self.RESULT_SKIP_VERSION))

        update_btn.SetDefault()
        update_btn.SetFocus()
        self.Bind(wx.EVT_CHAR_HOOK, self._on_key)
        self.Centre()

    def _on_key(self, event):
        if event.GetKeyCode() == wx.WXK_ESCAPE:
            self.EndModal(self.RESULT_LATER)
            return
        event.Skip()


# ─────────────────────────────────────────────────────────────
# 사용자 설명서 (Shift+F1)
# ─────────────────────────────────────────────────────────────

def _parse_manual_chapters(text: str) -> list[tuple[str, str]]:
    """manual.txt를 챕터 단위로 분할. 반환: [(제목, 본문), ...]

    챕터 경계: "숫자. 제목" 다음 줄이 ---- 또는 ==== 같은 구분선.
    구분선도 본문에 포함해 화면상 그대로 보이게 둔다.
    """
    lines = text.splitlines()
    # 챕터 시작 라인 인덱스 수집
    starts: list[tuple[int, str]] = []
    for i, line in enumerate(lines):
        m = re.match(r'^(\d+)\.\s+(.+)$', line.rstrip())
        if not m:
            continue
        # 다음 줄이 구분선(-, =) 연속 문자여야 챕터 헤더로 인정
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if nxt and set(nxt) <= {"-", "="} and len(nxt) >= 3:
                starts.append((i, line.rstrip()))

    if not starts:
        return []

    chapters: list[tuple[str, str]] = []
    for idx, (start, title) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        body = "\n".join(lines[start:end]).rstrip() + "\n"
        chapters.append((title, body))
    return chapters


class ManualDialog(wx.Dialog):
    """사용자 설명서 대화상자.

    좌측: 챕터 목록 (ListBox) — ↑/↓로 챕터 이동
    우측: 선택된 챕터 본문 (ReadOnly TextCtrl)
    Tab: 목록 ↔ 본문 포커스 전환, ESC: 닫기
    """

    def __init__(self, parent, chapters: list[tuple[str, str]], font_size: int):
        super().__init__(
            parent,
            title="초록멀티 사용자 설명서",
            size=(780, 560),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self._chapters = chapters

        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # 좌측: 챕터 목록
        left_sizer = wx.BoxSizer(wx.VERTICAL)
        list_label = wx.StaticText(panel, label="챕터 목록(&C):")
        titles = [t for (t, _) in chapters]
        self.list_box = wx.ListBox(
            panel, choices=titles, style=wx.LB_SINGLE, name="챕터 목록",
        )
        left_sizer.Add(list_label, 0, wx.ALL, 5)
        left_sizer.Add(self.list_box, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 우측: 본문
        right_sizer = wx.BoxSizer(wx.VERTICAL)
        body_label = wx.StaticText(panel, label="본문(&B):")
        self.body_text = wx.TextCtrl(
            panel, value="",
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_DONTWRAP | wx.HSCROLL,
            name="본문",
        )
        self.body_text.SetFont(make_font(font_size))
        right_sizer.Add(body_label, 0, wx.ALL, 5)
        right_sizer.Add(self.body_text, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        main_sizer.Add(left_sizer, 1, wx.EXPAND)
        main_sizer.Add(right_sizer, 2, wx.EXPAND)

        # 닫기 버튼
        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(main_sizer, 1, wx.EXPAND | wx.ALL, 5)
        close_btn = wx.Button(panel, wx.ID_CANCEL, "닫기(&X)")
        outer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(outer)

        # 이벤트 바인딩
        self.list_box.Bind(wx.EVT_LISTBOX, self._on_chapter_select)
        self.list_box.Bind(wx.EVT_LISTBOX_DCLICK, self._on_chapter_activate)
        self.list_box.Bind(wx.EVT_KEY_DOWN, self._on_list_key)
        self.body_text.Bind(wx.EVT_KEY_DOWN, self._on_body_key)

        # ESC로 닫기 — ID_CANCEL은 기본 바인딩됨
        self.SetEscapeId(wx.ID_CANCEL)

        # 초기 선택
        if chapters:
            self.list_box.SetSelection(0)
            self._show_chapter(0)
        self.list_box.SetFocus()
        self.Centre()

    def _show_chapter(self, idx: int):
        if 0 <= idx < len(self._chapters):
            title, body = self._chapters[idx]
            self.body_text.ChangeValue(body)
            self.body_text.SetInsertionPoint(0)

    def _on_chapter_select(self, event):
        self._show_chapter(self.list_box.GetSelection())

    def _on_chapter_activate(self, event):
        """더블클릭 또는 Enter → 본문으로 포커스"""
        idx = self.list_box.GetSelection()
        if idx != wx.NOT_FOUND:
            title, _ = self._chapters[idx]
            self.body_text.SetFocus()
            self.body_text.SetInsertionPoint(0)
            speak(f"{title} 본문")

    def _on_list_key(self, event: wx.KeyEvent):
        key = event.GetKeyCode()
        if key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
            self._on_chapter_activate(event)
            return
        if key == wx.WXK_TAB:
            self.body_text.SetFocus()
            return
        event.Skip()

    def _on_body_key(self, event: wx.KeyEvent):
        key = event.GetKeyCode()
        # Ctrl+PageUp / PageDown → 이전/다음 챕터
        if event.ControlDown() and key == wx.WXK_PAGEUP:
            self._move_chapter(-1)
            return
        if event.ControlDown() and key == wx.WXK_PAGEDOWN:
            self._move_chapter(+1)
            return
        if key == wx.WXK_TAB:
            self.list_box.SetFocus()
            return
        event.Skip()

    def _move_chapter(self, delta: int):
        cur = self.list_box.GetSelection()
        if cur == wx.NOT_FOUND:
            return
        new_idx = max(0, min(len(self._chapters) - 1, cur + delta))
        if new_idx == cur:
            return
        self.list_box.SetSelection(new_idx)
        self._show_chapter(new_idx)
        title = self._chapters[new_idx][0]
        speak(title)
