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

    def __init__(self, session: requests.Session,
                 current_user_id: str | None = None,
                 current_user_nickname: str | None = None,
                 current_user_rank: str | None = None):
        super().__init__(
            None,
            title=APP_NAME,
            size=(800, 600),
        )

        self.session = session
        # 현재 로그인한 사용자의 소리샘 아이디 (mb_id).
        # 게시물 수정/삭제 시 본인 여부 검증에 사용.
        self.current_user_id = current_user_id
        # 현재 로그인한 사용자의 닉네임.
        # 게시물 목록에 표시되는 작성자 닉네임(post.author)과 비교하여 본인
        # 게시물 여부를 빠르게 판단 (서버 HTTP 호출 없이).
        self.current_user_nickname = current_user_nickname
        # 초록등대 동호회 회원 등급. WriteDialog 의 공지 체크박스 노출 등에 사용.
        self.current_user_rank = current_user_rank
        self.menu_manager = MenuManager()
        self.menu_manager.load()

        # v1.7 — 즐겨찾기 매니저
        from bookmark_manager import BookmarkManager
        self.bookmark_manager = BookmarkManager()

        # v1.7 — 게시판 구독 매니저 (실 폴링은 _start_memo_notifier 시점에 시작)
        self._subscription_manager = None
        self._subscription_timer = None

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

        # 쪽지·메일 실시간 알림 폴링 먼저 시작 — 프로그램이 열리자마자 바로
        # 감지를 시작해야 한다. NAS 자동 연결보다 **앞에** 두어야 NAS 연결
        # 과정(또는 그 과정에서 뜰 수 있는 대화상자)이 메일/쪽지 첫 tick 을
        # 지연시키지 않는다.
        self._unread_memo_count = 0
        self._unread_mail_count = 0
        self._base_title = APP_NAME
        # 시작 시 제목 표시줄을 깨끗한 상태로 — 알림 첫 tick 결과로 갱신됨.
        self.SetTitle(self._base_title)
        wx.CallAfter(self._start_memo_notifier)

        # 알림 폴링이 자리잡은 뒤에 NAS 자동 마운트 시도 (저장된 자격증명이
        # 있을 때만, 백그라운드). 2초 정도 뒤에 시작하면 첫 알림 tick 이
        # 안전하게 먼저 돌고 사용자가 미확인 알림을 받을 수 있다.
        wx.CallLater(2000, self._try_auto_mount_nas)

        # 시작 시 자동 업데이트 확인 (설정에서 끌 수 있음). 로그인/메뉴 음성이
        # 먼저 끝나도록 몇 초 지연 후 백그라운드로 실행.
        wx.CallLater(3000, self._auto_update_check)

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
        """연결 성공 시 음성 안내만. 팝업은 띄우지 않는다.

        모달 팝업은 wx.Timer 기반 쪽지·메일 실시간 알림 tick 을 지연시키므로,
        자동 연결 성공 안내는 TTS 로만 전달하고 사용자가 계속 알림을 받을 수
        있도록 한다.
        """
        speak("초록등대 자료실에 연결되었습니다.")

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

        # v1.7 — 즐겨찾기 메뉴
        bookmarks_menu = wx.Menu()
        self.id_open_bookmarks = wx.NewIdRef()
        self.id_add_bookmark = wx.NewIdRef()
        self.id_command_palette = wx.NewIdRef()
        self.id_toggle_subscribe = wx.NewIdRef()
        self.id_open_subscriptions = wx.NewIdRef()
        self.id_convert_daisy = wx.NewIdRef()
        self.id_edit_templates = wx.NewIdRef()
        bookmarks_menu.Append(self.id_open_bookmarks, "즐겨찾기 열기(&O)\tCtrl+B")
        bookmarks_menu.Append(self.id_add_bookmark, "현재 위치를 즐겨찾기에 추가(&A)\tCtrl+D")
        bookmarks_menu.AppendSeparator()
        bookmarks_menu.Append(self.id_toggle_subscribe, "현재 게시판 구독 토글(&S)\tCtrl+Shift+S")
        bookmarks_menu.Append(self.id_open_subscriptions, "구독 목록 보기(&L)\tCtrl+Alt+L")
        bookmarks_menu.AppendSeparator()
        bookmarks_menu.Append(self.id_command_palette, "명령 도구 모음(&P)\tCtrl+P")
        menubar.Append(bookmarks_menu, "즐겨찾기(&K)")

        # 설정 메뉴
        settings_menu = wx.Menu()
        self.id_settings = wx.NewIdRef()
        self.id_theme_next = wx.NewIdRef()
        self.id_theme_prev = wx.NewIdRef()
        settings_menu.Append(self.id_settings, "설정(&T)\tF7")
        settings_menu.AppendSeparator()
        settings_menu.Append(self.id_theme_next, "다음 화면 테마(&N)\tF6")
        settings_menu.Append(self.id_theme_prev, "이전 화면 테마(&P)\tShift+F6")
        settings_menu.AppendSeparator()
        settings_menu.Append(self.id_download_dir, "다운로드 폴더 변경(&D)")
        self.id_edit_menu_file = wx.NewIdRef()
        self.id_reload_menu_file = wx.NewIdRef()
        self.id_reset_menu_file = wx.NewIdRef()
        settings_menu.AppendSeparator()
        settings_menu.Append(self.id_edit_menu_file, "메뉴 목록 파일 편집(&M)")
        settings_menu.Append(self.id_reload_menu_file, "메뉴 목록 파일 다시 읽기(&R)")
        settings_menu.Append(self.id_reset_menu_file, "메뉴 목록 자동 감지로 초기화(&I)")
        settings_menu.AppendSeparator()
        # v1.7 — 답장 템플릿 파일 편집
        settings_menu.Append(self.id_edit_templates, "답장 템플릿 파일 편집(&E)")
        menubar.Append(settings_menu, "설정(&S)")

        # 도구 메뉴
        tools_menu = wx.Menu()
        self.id_nas_connect = wx.NewIdRef()
        self.id_nas_logout = wx.NewIdRef()
        self.id_memo_inbox = wx.NewIdRef()
        self.id_memo_compose = wx.NewIdRef()
        self.id_mail_compose = wx.NewIdRef()
        tools_menu.Append(self.id_nas_connect, "초록등대 자료실 연결(&N)\tCtrl+N")
        tools_menu.Append(self.id_nas_logout, "초록등대 자료실 로그아웃(&O)")
        tools_menu.AppendSeparator()
        tools_menu.Append(self.id_memo_inbox, "쪽지함 열기(&M)\tCtrl+M")
        tools_menu.Append(self.id_memo_compose, "쪽지 쓰기\tCtrl+Shift+M")
        tools_menu.Append(self.id_mail_compose, "메일함 열기\tCtrl+Shift+E")
        self.id_memo_check_now = wx.NewIdRef()
        tools_menu.Append(self.id_memo_check_now, "알림 센터 열기\tCtrl+Shift+N")
        tools_menu.AppendSeparator()
        # v1.7 — DAISY 도서 변환은 도구 메뉴에 위치
        tools_menu.Append(self.id_convert_daisy, "DAISY 도서 변환(&D)\tCtrl+Alt+D")
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
        self.Bind(wx.EVT_MENU, self._on_menu_nas_logout, id=self.id_nas_logout)
        self.Bind(wx.EVT_MENU, self.on_open_memo_inbox, id=self.id_memo_inbox)
        self.Bind(wx.EVT_MENU, self.on_open_memo_compose, id=self.id_memo_compose)
        self.Bind(wx.EVT_MENU, self.on_open_mail_compose, id=self.id_mail_compose)
        self.Bind(wx.EVT_MENU, self.on_memo_check_now, id=self.id_memo_check_now)
        self.Bind(wx.EVT_MENU, self.on_edit_menu_file, id=self.id_edit_menu_file)
        self.Bind(wx.EVT_MENU, self.on_reload_menu_file, id=self.id_reload_menu_file)
        self.Bind(wx.EVT_MENU, self.on_reset_menu_file, id=self.id_reset_menu_file)
        self.Bind(wx.EVT_MENU, self.on_open_bookmarks, id=self.id_open_bookmarks)
        self.Bind(wx.EVT_MENU, self.on_add_bookmark, id=self.id_add_bookmark)
        self.Bind(wx.EVT_MENU, self.on_open_command_palette, id=self.id_command_palette)
        self.Bind(wx.EVT_MENU, self.on_toggle_subscription, id=self.id_toggle_subscribe)
        self.Bind(wx.EVT_MENU, self.on_open_subscriptions, id=self.id_open_subscriptions)
        self.Bind(wx.EVT_MENU, self.on_convert_daisy, id=self.id_convert_daisy)
        self.Bind(wx.EVT_MENU, self.on_edit_reply_templates, id=self.id_edit_templates)
        self.Bind(wx.EVT_MENU, self.on_theme_next, id=self.id_theme_next)
        self.Bind(wx.EVT_MENU, self.on_theme_prev, id=self.id_theme_prev)

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

        # 상황별 팝업(컨텍스트) 메뉴 — 오른쪽 클릭 / Menu 키 / Shift+F10
        self.textctrl.Bind(wx.EVT_CONTEXT_MENU, self._on_textctrl_context_menu)

    def _on_textctrl_context_menu(self, event):
        """현재 화면(current_view)에 맞는 액션만 모은 팝업 메뉴 표시.

        이동/페이지 탐색 같은 기본 내비게이션 항목은 의도적으로 제외.
        """
        if self.current_view != VIEW_POST_LIST:
            # 메인/하위 메뉴에서는 상황별 액션이 거의 없음 → 팝업 표시 안 함
            return

        menu = wx.Menu()
        menu.Append(self.id_post_write, "게시물 작성(&W)\tW")
        menu.Append(self.id_post_edit, "게시물 수정(&M)\tAlt+M")
        menu.Append(self.id_post_delete, "게시물 삭제(&D)\tAlt+D")
        menu.AppendSeparator()
        # v1.7 — 게시물을 열지 않고 첨부파일 즉시 다운로드 (D 단축키와 동일)
        id_dl = wx.NewIdRef()
        menu.Append(id_dl, "선택한 게시물 첨부파일 저장(&S)\tD")
        self.Bind(
            wx.EVT_MENU,
            lambda e: self._download_post_attachments_from_list(),
            id=id_dl,
        )
        menu.AppendSeparator()
        menu.Append(self.id_search, "게시물 검색(&F)\tCtrl+F")
        menu.Append(self.id_board_refresh, "게시판 새로고침(&R)\tF5")

        self.PopupMenu(menu)
        menu.Destroy()

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
            # v1.7 — DAISY 변환 / 구독 목록 보기
            wx.AcceleratorEntry(
                wx.ACCEL_CTRL | wx.ACCEL_ALT, ord("D"), self.id_convert_daisy,
            ),
            wx.AcceleratorEntry(
                wx.ACCEL_CTRL | wx.ACCEL_ALT, ord("L"), self.id_open_subscriptions,
            ),
            # 화면 테마 순환 — F6 (다음) / Shift+F6 (이전)
            wx.AcceleratorEntry(
                wx.ACCEL_NORMAL, wx.WXK_F6, self.id_theme_next,
            ),
            wx.AcceleratorEntry(
                wx.ACCEL_SHIFT, wx.WXK_F6, self.id_theme_prev,
            ),
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

    # ── 화면 테마 순환 (F6 / Shift+F6) ──

    def _cycle_theme(self, direction: int):
        """저장된 THEME_PRESETS 키 목록을 순환해 +/-1 위치로 이동."""
        try:
            from theme import (
                THEME_PRESETS, load_theme_key, set_current_theme,
                get_current_theme_name,
            )
        except Exception:
            return
        keys = list(THEME_PRESETS.keys())
        if not keys:
            return
        cur_key = load_theme_key()
        try:
            idx = keys.index(cur_key)
        except ValueError:
            idx = 0
        new_key = keys[(idx + direction) % len(keys)]
        set_current_theme(new_key)
        self._apply_full_theme()
        try:
            speak(f"테마 변경: {get_current_theme_name()}")
        except Exception:
            pass

    def on_theme_next(self, event=None):
        """F6: 다음 화면 테마로 변경."""
        self._cycle_theme(1)

    def on_theme_prev(self, event=None):
        """Shift+F6: 이전 화면 테마로 변경."""
        self._cycle_theme(-1)

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
        # 메인 메뉴를 그릴 때마다 자료실·엔터테인먼트 자료실 보장 로직 재실행.
        # 어떤 이유로든 목록에서 빠지거나 순서가 틀어진 경우를 화면 표시 직전에
        # 복원한다.
        try:
            if self.menu_manager._ensure_forced_club_menus():
                self.menu_manager.save()
        except Exception:
            pass
        menu_names = self.menu_manager.get_display_names()
        self._update_textctrl(menu_names, "메뉴 목록")
        self.status_bar.SetStatusText("준비", 0)
        try:
            from sound import play_event
            play_event("main_menu_return")
        except Exception:
            pass

    def _show_sub_menu(self, sub_menus: list[SubMenuItem], menu_name: str,
                       base_url: str = ""):
        self.current_view = VIEW_SUB_MENU
        self.current_menu_name = menu_name
        self.SetTitle(f"{APP_NAME} - {menu_name}")
        # 현재 하위 메뉴 페이지의 소스 URL. 뒤따르는 필터가 cl=<code> 를 추출해
        # 다른 클럽 카테고리 링크를 걸러내는 데 사용.
        self.current_sub_menu_url = base_url

        clean_menu = re.sub(r'^\d+[\.\)]\s*', '', menu_name).strip() if menu_name else ""

        # 카테고리 헤더로 제거할 텍스트 목록
        header_noise = {
            "홈", "home",
            "글쓰기", "게시판관리", "멀티업로드",
            "img", "관리자", "철머",
            "로그아웃", "돌아가기",
            "소리샘 동사무소", "동사무소",
            # 초록등대 자료실(cl=green4) / 엔터테인먼트 자료실(cl=green6) 진입 시
            # 자동 노출되는 클럽 네비게이션 링크 — 사용자에겐 노이즈.
            "일반 동호회 바로가기", "일반동호회 바로가기",
            "초록등대 바로가기", "초록등대",
        }
        # 현재 메뉴명만 노이즈로 추가
        if clean_menu:
            header_noise.add(clean_menu)

        # 메인 메뉴 URL / 바로가기 코드 집합 — 상위 내비게이션 식별 보조용.
        from menu_manager import extract_shortcut_code

        main_menu_urls = {"/", ""}
        main_menu_codes: set[str] = set()
        for mi in self.menu_manager.menus:
            main_menu_urls.add(mi.url)
            code = extract_shortcut_code(mi.url)
            if code:
                main_menu_codes.add(code.strip().lower())

        # 현재 보고 있는 페이지의 컨텍스트 코드. 우선순위:
        # 1) 호출자가 넘긴 base_url  2) current_board_url  3) 비어있음
        source_url = (
            getattr(self, "current_sub_menu_url", "") or self.current_board_url or ""
        )
        current_cl = ""
        m_cur = re.search(r"[?&]cl=([^&#]+)", source_url)
        if m_cur:
            current_cl = m_cur.group(1).strip().lower()
        current_code = (
            current_cl
            or (extract_shortcut_code(source_url) or "").strip().lower()
        )

        # 클럽(ar.club)·현재 컨텍스트가 명확할 때 "화이트리스트 + 명시 거부"
        # 방식으로 필터링. 블랙리스트(메인 메뉴 코드와 비교) 는 메인 메뉴에
        # 없는 카테고리(예: circle)를 놓치기 때문에 휴리스틱을 추가한다.
        strict_scope = bool(current_cl) and "ar.club" in source_url

        # 진단 로그: 현재 하위 메뉴 필터링 결과를 파일로 남겨 어떤 URL 이 왜
        # 유지/제외되는지 확인할 수 있게 한다. 사용자가 문제 상황에서 이 파일을
        # 제공하면 필터 규칙을 정확히 맞출 수 있다.
        from config import DATA_DIR
        _diag_path = os.path.join(DATA_DIR, "submenu_filter.log")
        _diag_lines = [
            "=== _show_sub_menu 진단 ===",
            f"menu_name     = {menu_name}",
            f"base_url      = {base_url}",
            f"source_url    = {source_url}",
            f"current_cl    = {current_cl}",
            f"current_code  = {current_code}",
            f"strict_scope  = {strict_scope}",
            f"main_menu_codes = {sorted(main_menu_codes)}",
            f"raw sub_menus 개수 = {len(sub_menus)}",
            "--- 전체 후보 목록 ---",
        ]
        for _i, _m in enumerate(sub_menus):
            _diag_lines.append(f"  [{_i}] url={_m.url!r}  text={_m.name!r}")
        _diag_lines.append("--- 필터 결과 ---")

        # 필터링된 하위메뉴와 표시 항목을 동기화
        filtered_subs = []
        display_items = ["0. 메인 메뉴로 돌아가기"]
        seen_texts = set()
        num = 1
        for m in sub_menus:
            url = (m.url or "").strip()
            url_lower = url.lower()

            # 게시글 본문 링크(wr_id=) 는 하위 메뉴가 아님 — 제외
            if "wr_id=" in url_lower:
                _diag_lines.append(
                    f"  REJECT [post_link] url={url!r} text={m.name!r}"
                )
                continue

            # 부모 카테고리/자기 자신으로 돌아가는 navigation 링크 거부.
            # sorisem 은 클럽 내부 카테고리 페이지(/?mo=greenN&cl=green 등) 응답에
            # "일반 동호회"(/?mo=circle), 부모 클럽 자체(/?mo=green&cl=green) 같은
            # 상위 navigation 링크를 sub-menu 처럼 끼워 넣는다. sub-menu 노이즈로 거부.
            PARENT_NAV_URLS_LOW = {
                "/?mo=circle",
                "/?mo=circle&cl=circle",
                "/?mo=potion",
                "/?mo=potion&cl=potion",
            }
            if url_lower in PARENT_NAV_URLS_LOW:
                _diag_lines.append(
                    f"  REJECT [parent_nav] url={url!r} text={m.name!r}"
                )
                continue

            # 현재 클럽 자체로 돌아가는 self-link 거부 (부모 클럽 hub).
            # 예: cl=green 컨텍스트에서 /?mo=green&cl=green 또는
            # /plugin/ar.club/?cl=green 은 자기 자신 링크 → sub-menu 노이즈.
            if current_cl:
                self_urls_low = {
                    f"/?mo={current_cl}&cl={current_cl}".lower(),
                    f"/plugin/ar.club/?cl={current_cl}".lower(),
                }
                if url_lower in self_urls_low:
                    _diag_lines.append(
                        f"  REJECT [self_nav] url={url!r} text={m.name!r}"
                    )
                    continue

            # 초록등대 동호회 자료실/엔터테인먼트 자료실 링크는 초록등대 동호회
            # 컨텍스트(cl=green)에서 하위 메뉴를 볼 때만 표시. 소리샘 자료실
            # (mo=pds) 같은 다른 컨텍스트의 하위 메뉴에서는 엉뚱하게 끼어들지
            # 않도록 제외.
            if url in ("/plugin/ar.club/?cl=green4", "/plugin/ar.club/?cl=green6"):
                if current_cl != "green":
                    continue

            # 브레드크럼(경로 안내) 링크 제거: 메인 메뉴 URL과 동일한 항목
            # 단, 자료실·엔터테인먼트 자료실은 초록등대 동호회 컨텍스트
            # (cl=green)에서 하위 메뉴를 볼 때만 예외적으로 표시.
            if url in main_menu_urls:
                from menu_manager import _forced_shortcut_code
                if not (current_cl == "green" and _forced_shortcut_code(m.name)):
                    continue

            sub_code = (extract_shortcut_code(url) or "").strip().lower()

            if strict_scope:
                _reject_reason = None
                has_mo = bool(re.search(r"[?&]mo=", url_lower))

                # ─── mo= 취급: 단독이면 최상위, cl= 동반이면 관계 기반 판정 ───
                if has_mo:
                    _m_cl_in_mo = re.search(r"[?&]cl=([^&#]+)", url_lower)
                    if not _m_cl_in_mo:
                        _reject_reason = "R1 mo= top-level (no cl=)"
                    else:
                        _cl_in_mo = _m_cl_in_mo.group(1).strip().lower()
                        _common_mo = 0
                        for c1, c2 in zip(_cl_in_mo, current_cl):
                            if c1 == c2:
                                _common_mo += 1
                            else:
                                break
                        if _cl_in_mo != current_cl and _common_mo < 3:
                            _reject_reason = (
                                f"R1 mo= unrelated cl={_cl_in_mo}"
                            )
                        # else: 동일/관련 클럽의 섹션 → 통과
                elif re.search(r"[?&]clp=", url_lower):
                    _reject_reason = "R1 clp="
                elif url_lower.startswith(("http://", "https://")):
                    _reject_reason = "R2 external"
                elif url_lower in ("/", "") or url_lower.startswith("/mypage"):
                    _reject_reason = "R3 home/mypage"

                # ─── 클럽 코드 비교 (부모/형제/무관 판별) ───
                # mo= 가 함께 있는 부모-클럽 링크는 "부모의 특정 섹션" 이므로
                # breadcrumb 으로 간주해 거부하지 않는다.
                if _reject_reason is None:
                    m_cl_any = re.search(r"[?&]cl=([^&#]+)", url_lower)
                    m_path_club = re.search(
                        r"/plugin/ar\.club/([a-zA-Z0-9_]+)(?:/|\?|$)", url_lower,
                    )
                    link_cl = None
                    if m_cl_any:
                        link_cl = m_cl_any.group(1).strip().lower()
                    elif m_path_club:
                        link_cl = m_path_club.group(1).lower()

                    if link_cl and link_cl != current_cl:
                        common_len = 0
                        for c1, c2 in zip(link_cl, current_cl):
                            if c1 == c2:
                                common_len += 1
                            else:
                                break
                        if common_len < 3:
                            _reject_reason = (
                                f"R-club unrelated (link_cl={link_cl}, "
                                f"common={common_len})"
                            )
                        elif current_cl.startswith(link_cl) and not has_mo:
                            _reject_reason = (
                                f"R-club parent breadcrumb (link_cl={link_cl})"
                            )

                    if (
                        _reject_reason is None
                        and "/plugin/ar.club/" in url_lower
                        and "cl=" not in url_lower
                        and not m_path_club
                    ):
                        _reject_reason = "R4 ar.club listing"

                if _reject_reason is not None:
                    _diag_lines.append(
                        f"  REJECT [{_reject_reason}] url={url!r} text={m.name!r}"
                    )
                    continue
            else:
                # 클럽 아닌 일반 컨텍스트:
                # sorisem.net 은 `/?mo=XXX` 로 중첩 카테고리를 표현한다.
                # 예) /?mo=potion(동호회) → /?mo=circle(일반 동호회) → /?mo=...
                # 따라서 mo= 값이 현재와 다르다고 무조건 거부하면 안 된다.
                # main_menu_urls 검사(이전 단계)가 이미 메인 메뉴와 정확히 같은
                # URL 을 거부했으므로 여기서는 추가 거부를 최소화한다.
                if re.search(r"[?&]clp=", url_lower):
                    continue
                # 다른 메인 메뉴 코드와 일치하는 경우 거부.
                # 단, 자료실·엔터테인먼트 자료실 예외 + main_menu_urls 검사를
                # 이미 통과한 항목이므로 여기서 거부되는 것은 동일 코드를 가진
                # 다른 URL 패턴(드문 경우) 정도로 한정.
                from menu_manager import _forced_shortcut_code
                if sub_code and sub_code in main_menu_codes and sub_code != current_code:
                    if not (current_cl == "green" and _forced_shortcut_code(m.name)):
                        continue

            original_text = m.display_text
            text = original_text

            # 상위 메뉴명 접두사 제거
            if clean_menu and text.startswith(clean_menu):
                text = text[len(clean_menu):].lstrip(" ·:>-")
            elif menu_name and text.startswith(menu_name):
                text = text[len(menu_name):].lstrip(" ·:>-")

            # 기존 번호 제거
            text = re.sub(r'^\d+[\.\)]\s*', '', text).strip()

            # 카테고리 헤더 / 노이즈 제거.
            # 비교 전 텍스트를 정규화 — NBSP/Tab/연속 공백을 단일 공백으로 합쳐
            # 사이트 응답에 미세한 공백 차이가 있어도 매칭되도록 한다.
            text_norm = re.sub(r"\s+", " ", text).strip()
            if text_norm.lower() in {
                re.sub(r"\s+", " ", h).strip().lower() for h in header_noise
            }:
                _diag_lines.append(
                    f"  REJECT [header_noise] url={url!r} text={original_text!r}"
                )
                continue
            # "동사무소" 가 포함된 텍스트(예: "소리샘 동사무소")는 모든 페이지에서
            # 노이즈로 간주
            if "동사무소" in text:
                _diag_lines.append(
                    f"  REJECT [contains_동사무소] url={url!r} text={original_text!r}"
                )
                continue
            # 4. 자료실(cl=green4) / 6. 엔터테인먼트 자료실(cl=green6) 진입 시
            # 자동 노출되는 클럽 네비게이션 — 부분 문자열 매칭으로 강하게 거부.
            # 사용자가 미세한 공백·구두점 차이로 매번 패치를 요청하지 않도록 한다.
            if (
                "동호회 바로가기" in text_norm
                or "동호회바로가기" in text_norm
                or "초록등대 바로가기" in text_norm
                or "초록등대바로가기" in text_norm
            ):
                _diag_lines.append(
                    f"  REJECT [club_nav_link] url={url!r} text={original_text!r}"
                )
                continue
            if not text:
                _diag_lines.append(
                    f"  REJECT [text_empty] url={url!r} text={original_text!r}"
                )
                continue
            # 중복 제거
            if text.lower() in seen_texts:
                _diag_lines.append(
                    f"  REJECT [duplicate_text] url={url!r} text={original_text!r}"
                )
                continue
            seen_texts.add(text.lower())

            # 바로가기 코드 — 이름이 "자료실"/"엔터테인먼트 자료실" 이고
            # 현재 컨텍스트가 초록등대 동호회(cl=green)일 때만 강제로
            # green4/green6 표시. 그 외에는 URL 기반.
            from menu_manager import extract_shortcut_code, _forced_shortcut_code
            forced = _forced_shortcut_code(text) if current_cl == "green" else ""
            if forced:
                code = forced
            else:
                code = extract_shortcut_code(m.url)
            if code:
                display_items.append(f"{num}. {text} (바로가기 코드: {code})")
            else:
                display_items.append(f"{num}. {text}")
            filtered_subs.append(m)
            _diag_lines.append(
                f"  KEEP    url={url!r} text={original_text!r} code={code!r}"
            )
            num += 1

        # 필터가 모든 항목을 거부했지만 원본 후보 목록은 있을 때:
        # 사용자가 빈 화면을 보지 않도록, 자기 자신·헤더 노이즈만 걸러내고
        # 나머지를 그대로 보여준다. (예: /?mo=prg, /?mo=potion 같은 카테고리
        # 페이지에서 sub-link 들이 main_menu_codes 와 겹쳐 잘못 잘려나가는 경우)
        if not filtered_subs and sub_menus:
            display_items = ["0. 메인 메뉴로 돌아가기"]
            seen_texts2: set[str] = set()
            num2 = 1
            from menu_manager import extract_shortcut_code as _esc
            # rescue: 우선 메인 메뉴와 *겹치지 않는* 항목만 모은다. 대부분의
            # 카테고리 페이지는 사이드바(메인 메뉴) 와 그 페이지 고유 항목을
            # 함께 갖는다. 메인 메뉴와 겹치는 부분은 사이드바이고, 나머지가
            # 진짜 하위 메뉴.
            # 진짜 하위 메뉴가 하나도 없으면(=페이지가 메인 사이드바만 돌려준
            # 상황: /?mo=pds, /?mo=lib2013 같은 빈/이름만 있는 카테고리)
            # 사용자가 헷갈리지 않도록 안내 문구만 보여 준다 (메인 메뉴 중복
            # 항목을 다시 노출하지 않는다).
            rescue_candidates: list = []
            had_only_main_menu = True
            for m in sub_menus:
                url = (m.url or "").strip()
                if not url or url in ("/", "", "#") or url == source_url:
                    continue
                if "wr_id=" in url.lower():
                    continue
                text = (m.display_text or "").strip()
                text = re.sub(r"^\d+[\.\)]\s*", "", text).strip()
                if not text or len(text) < 2 or len(text) > 60:
                    continue
                text_norm_r = re.sub(r"\s+", " ", text).strip()
                if text_norm_r.lower() in {
                    re.sub(r"\s+", " ", h).strip().lower() for h in header_noise
                }:
                    continue
                if "동사무소" in text:
                    continue
                if (
                    "동호회 바로가기" in text_norm_r
                    or "동호회바로가기" in text_norm_r
                    or "초록등대 바로가기" in text_norm_r
                    or "초록등대바로가기" in text_norm_r
                ):
                    continue
                if text.lower() in seen_texts2:
                    continue
                seen_texts2.add(text.lower())
                # 메인 메뉴 URL/코드와 정확히 겹치는 항목은 별도로 두고, 진짜
                # 하위 메뉴 후보(rescue_candidates) 와 분리.
                sub_code_r = (extract_shortcut_code(url) or "").strip().lower()
                is_main = (
                    url in main_menu_urls
                    or (sub_code_r and sub_code_r in main_menu_codes
                        and sub_code_r != current_code)
                )
                if not is_main:
                    rescue_candidates.append((url, text, m))
                    had_only_main_menu = False

            for url, text, m in rescue_candidates:
                code = _esc(url) or ""
                if code:
                    display_items.append(f"{num2}. {text} (바로가기 코드: {code})")
                else:
                    display_items.append(f"{num2}. {text}")
                filtered_subs.append(m)
                _diag_lines.append(
                    f"  RESCUE  url={url!r} text={text!r}"
                )
                num2 += 1

        if not filtered_subs:
            # 페이지가 비어있다는 것 + URL 이 잘못되었을 가능성을 동시에 안내.
            display_items = [
                "0. 메인 메뉴로 돌아가기",
                f"이 페이지({source_url})는 별도 하위 메뉴를 제공하지 않습니다.",
                "사이트에서 다른 URL을 사용한다면 설정 메뉴 > '메뉴 목록 파일 편집' 에서 URL을 바꿔 주세요.",
            ]

        # 진단 로그 파일 저장
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(_diag_path, "w", encoding="utf-8") as _f:
                _f.write("\n".join(_diag_lines))
                _f.write(f"\n\n최종 표시 개수: {len(filtered_subs)}\n")
        except Exception:
            pass

        self.current_sub_menus = filtered_subs
        self._update_textctrl(display_items, f"{menu_name} 하위 메뉴")
        self.status_bar.SetStatusText(f"{menu_name} - {len(filtered_subs)}개 하위 메뉴", 0)

    def _show_post_list(self, posts: list[PostItem], menu_name: str,
                        board_url: str = "", page: int = 1):
        self.current_view = VIEW_POST_LIST
        # 4. 자료실(cl=green4) / 6. 엔터테인먼트 자료실(cl=green6) 같은 페이지는
        # sorisem 응답에 "일반 동호회 바로가기", "초록등대 바로가기" 같은 클럽
        # 네비게이션 링크가 게시물처럼 섞여 들어오는 사례가 있다. 게시물 목록
        # 표시 직전에 제목 기반으로 한 번 더 거른다.
        def _is_nav_noise(p) -> bool:
            t = re.sub(r"\s+", " ", (p.title or "")).strip()
            return (
                "동호회 바로가기" in t
                or "동호회바로가기" in t
                or "초록등대 바로가기" in t
                or "초록등대바로가기" in t
            )
        posts = [p for p in posts if not _is_nav_noise(p)]
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
            dialog = PostDialog(
                self, content, self.session,
                current_user_id=self.current_user_id,
                current_user_nickname=self.current_user_nickname,
            )
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
                # cl= 가 있는 board.php 호출은 sorisem 이 hub 컨텍스트를 요구한다.
                # worker 내부에서 hub 를 먼저 동기적으로 호출해 세션 컨텍스트를 갖춘
                # 뒤 본 요청을 보낸다. 한 세션당 hub 별로 한 번만 워밍업.
                referer = None
                if "bo_table=" in url and re.search(r"[?&]cl=([^&#]+)", url):
                    m_cl = re.search(r"[?&]cl=([^&#]+)", url)
                    cl_val = m_cl.group(1) if m_cl else ""
                    if cl_val:
                        # 매 호출마다 hub 를 재방문해 세션 컨텍스트를 강제로 갱신
                        # (캐시 사용 금지). sorisem 은 hub fetch 한 번 후 다른
                        # 카테고리 hub 를 거치면 cl 컨텍스트가 풀리는 동작을 보여
                        # 캐싱하면 권한 오류로 이어진다.
                        hub_path = f"/?mo={cl_val}&cl={cl_val}"
                        try:
                            hub_resp = self.session.get(
                                f"{SORISEM_BASE_URL}{hub_path}", timeout=15,
                            )
                            try:
                                from config import DATA_DIR
                                import os as _os
                                safe = re.sub(r"[^A-Za-z0-9]+", "_", hub_path)[:40]
                                _os.makedirs(DATA_DIR, exist_ok=True)
                                with open(
                                    _os.path.join(DATA_DIR, f"warmup_{safe}.html"),
                                    "w", encoding="utf-8",
                                ) as _wf:
                                    _wf.write(hub_resp.text)
                            except Exception:
                                pass
                        except Exception:
                            pass
                        referer = f"{SORISEM_BASE_URL}{hub_path}"

                # ar.club nested 클럽 (예: cl=hims, cl=green3) 직접 호출 시
                # 부모 클럽 hub (cl=green) 를 거치지 않으면 sorisem 이
                # "게시판 접근권한이 없습니다" 로 거부한다. 메인 메뉴에서 클릭
                # 하면 자연스럽게 cl=green 을 먼저 거치지만, 바로가기 코드로
                # 직접 진입하는 경우엔 부모 클럽을 명시적으로 워밍업해야 한다.
                if "/plugin/ar.club/" in url:
                    m_cl_club = re.search(r"[?&]cl=([^&#]+)", url)
                    cl_club = m_cl_club.group(1).lower() if m_cl_club else ""
                    # cl=green 자체는 부모 — 워밍업 불필요.
                    # 그 외(cl=hims, cl=green2 등) 는 cl=green 부모 hub 를 먼저 호출.
                    if cl_club and cl_club != "green":
                        parent_hub = "/plugin/ar.club/?cl=green"
                        try:
                            self.session.get(
                                f"{SORISEM_BASE_URL}{parent_hub}", timeout=15,
                            )
                        except Exception:
                            pass
                        if not referer:
                            referer = f"{SORISEM_BASE_URL}{parent_hub}"

                headers = {"Referer": referer} if referer else None
                resp = self.session.get(full_url, timeout=15, headers=headers)
                html = resp.text

                # 진단용 세션 상태 로그.
                try:
                    from config import DATA_DIR
                    import os as _os
                    _os.makedirs(DATA_DIR, exist_ok=True)
                    with open(
                        _os.path.join(DATA_DIR, "session_debug.log"),
                        "a", encoding="utf-8",
                    ) as _sf:
                        cookies_str = "; ".join(
                            f"{c.name}={c.value[:20]}..."
                            if len(c.value) > 20
                            else f"{c.name}={c.value}"
                            for c in self.session.cookies
                        )
                        member_match = re.search(
                            r'g5_is_member\s*=\s*"([^"]*)"', html,
                        )
                        member_val = member_match.group(1) if member_match else "?"
                        _sf.write(
                            f"\n=== {url} ===\n"
                            f"status={resp.status_code} bytes={len(html)}\n"
                            f"g5_is_member={member_val!r}\n"
                            f"referer={referer}\n"
                            f"cookies={cookies_str}\n"
                        )
                except Exception:
                    pass

                # 세션 만료 / 권한 거부 자동 복구. 응답에 sorisem 의 표준 거부
                # 메시지가 보이면 저장된 자격증명으로 재로그인 후 한 번 더 시도.
                access_denied = (
                    "접근권한이 없습니다" in html
                    or "오류안내" in html
                )
                if access_denied and self._try_relogin():
                    try:
                        # 재로그인 후엔 hub 를 다시 방문해 컨텍스트를 다시 만든다.
                        if referer:
                            self.session.get(referer, timeout=15)
                        resp2 = self.session.get(full_url, timeout=15, headers=headers)
                        html = resp2.text
                        try:
                            from config import DATA_DIR
                            import os as _os
                            with open(
                                _os.path.join(DATA_DIR, "session_debug.log"),
                                "a", encoding="utf-8",
                            ) as _sf:
                                _sf.write(
                                    f"--- after relogin retry ---\n"
                                    f"status={resp2.status_code} bytes={len(html)}\n"
                                )
                        except Exception:
                            pass
                    except Exception:
                        pass

                wx.CallAfter(callback, html, None)
            except requests.exceptions.RequestException as e:
                wx.CallAfter(callback, None, str(e))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _try_relogin(self) -> bool:
        """저장된 자격증명으로 sorisem 에 다시 로그인.

        반환값: 재로그인 성공 시 True. 실패하거나 자격증명이 저장되지 않았으면 False.
        한 번 시도하면 30초 동안 재시도하지 않는다 (실패 폭주 방지).
        """
        import time
        last = getattr(self, "_last_relogin_attempt", 0)
        if time.time() - last < 30:
            return False
        self._last_relogin_attempt = time.time()

        try:
            from credentials import load_credentials
            from authenticator import Authenticator
            creds = load_credentials()
            if not creds:
                return False
            user_id, password = creds
            # 같은 세션 객체를 재사용해 쿠키를 그 자리에서 갱신.
            auth = Authenticator()
            auth.session = self.session
            result = auth._login(user_id, password)
            return bool(result and result.is_success)
        except Exception:
            return False

    def _warmup_session_for(self, url: str):
        """가상 하위 메뉴 진입 시 sorisem 세션 컨텍스트를 미리 설정.

        sorisem 은 `/?mo=XXX&cl=XXX` hub 페이지를 한 번 거쳐야 해당 카테고리의
        board.php 게시판 호출이 허용된다. 가상 하위 메뉴는 hub fetch 를 생략하므로
        백그라운드 스레드에서 hub URL 과 (가능하면) cl 컨텍스트 URL 을 GET 해
        세션 쿠키를 채워둔다. 결과는 사용하지 않고 폐기.
        """
        warm_urls = []
        # url 자체가 hub 형태(`?mo=XXX&cl=XXX` 또는 `?mo=XXX`)면 그대로 GET.
        if url and "/?mo=" in url:
            warm_urls.append(url)
        # cl 값을 추출해 `/?mo=cl&cl=cl` 형태도 같이 워밍업
        if url:
            m = re.search(r"[?&]cl=([^&#]+)", url)
            if m:
                cl_val = m.group(1)
                hub2 = f"/?mo={cl_val}&cl={cl_val}"
                if hub2 not in warm_urls:
                    warm_urls.append(hub2)

        if not warm_urls:
            return

        seen = getattr(self, "_warmed_session_urls", None)
        if seen is None:
            seen = set()
            self._warmed_session_urls = seen

        def worker():
            for u in warm_urls:
                if u in seen:
                    continue
                try:
                    full = u if u.startswith("http") else f"{SORISEM_BASE_URL}{u}"
                    self.session.get(full, timeout=15)
                    seen.add(u)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _load_and_show(self, url: str, name: str):
        # v1.7 — 가상 하위 메뉴 처리. sorisem 의 hub 페이지가 비어 있거나
        # 별도로 sub 목록을 응답하지 않는 메인 메뉴 항목(예: 7. 전자도서관) 은
        # 코드에서 정의한 sub-item 목록을 그대로 표시한다. 네트워크 fetch 없이
        # 즉시 sub-menu 화면으로 전환.
        try:
            from menu_manager import VIRTUAL_SUBMENUS
            virt = VIRTUAL_SUBMENUS.get(url)
        except Exception:
            virt = None
        if virt:
            self.status_bar.SetStatusText(f"{name} 로딩 중...", 0)
            speak(f"{name} 로딩 중입니다.")
            # 세션 컨텍스트 워밍업 — sorisem 은 /?mo=XXX&cl=XXX hub 를 거치지
            # 않은 채 cl=XXX 게시판을 직접 호출하면 "게시판 접근권한이 없습니다"
            # 로 거부한다. 가상 하위 메뉴 진입 시에는 hub fetch 가 생략되므로
            # 여기서 백그라운드로 한 번 GET 해 세션 쿠키를 채워둔다.
            self._warmup_session_for(url)
            virtual_items = [
                SubMenuItem(name=sub_name, url=sub_url)
                for sub_name, sub_url, _t in virt
            ]
            self._show_sub_menu(virtual_items, name, base_url=url)
            self.status_bar.SetStatusText("준비", 0)
            return

        self.status_bar.SetStatusText(f"{name} 로딩 중...", 0)
        speak(f"{name} 로딩 중입니다.")

        board_url = url

        def on_loaded(html, error):
            # board_url 을 콜백 안에서 자동 폴백 결과로 갱신할 수 있도록
            # 외부 클로저 변수를 명시적으로 끌어 쓴다 (없으면 로컬로 처리되어
            # 같은 이름을 먼저 읽는 줄에서 UnboundLocalError 발생).
            nonlocal board_url
            if error:
                speak(f"페이지를 불러올 수 없습니다. {error}")
                self.status_bar.SetStatusText("준비", 0)
                return

            if not html or len(html) < 50:
                speak("빈 응답을 받았습니다.")
                self.status_bar.SetStatusText("준비", 0)
                return

            # 카테고리 랜딩 페이지(`/?mo=XXX` 형식, bo_table 없음) 판정.
            # cl= 가 같이 있더라도 bo_table 만 없으면 카테고리 페이지로 간주한다.
            # (예: 8.노원시각장애인학습지원센터 = /?mo=edu2013&cl=edu2013 — 클럽
            # 메인 페이지지만 본문이 아니라 하위 섹션 모음이므로 동일 흐름 사용.)
            is_category_page = (
                bool(re.search(r"[?&]mo=[a-zA-Z0-9_]+", board_url))
                and "bo_table=" not in board_url
            )

            # 진단용: bo_table 이 없는 페이지(/?mo=... 또는 /?mo=...&cl=...) 는
            # 항상 HTML 을 data/ 로 덤프해 사용자가 공유할 수 있게 한다.
            if "bo_table=" not in board_url:
                try:
                    from config import DATA_DIR
                    safe_dbg = re.sub(r"[^A-Za-z0-9]+", "_", board_url)[:40]
                    os.makedirs(DATA_DIR, exist_ok=True)
                    with open(
                        os.path.join(DATA_DIR, f"category_{safe_dbg}.html"),
                        "w", encoding="utf-8",
                    ) as _df:
                        _df.write(html)
                except Exception:
                    pass

            if is_category_page:
                sub_menus = parse_sub_menus(html, base_url=board_url)
                # parse_sub_menus 가 1~2개만 반환해도 fallback 을 합쳐 더 넓게.
                fallback_links = self._extract_fallback_links(html)
                # 합치되 중복 제거 (URL 기준)
                merged: list = []
                seen_urls: set[str] = set()
                for item in list(sub_menus or []) + list(fallback_links or []):
                    if item.url in seen_urls:
                        continue
                    seen_urls.add(item.url)
                    merged.append(item)

                # v1.7 — 자동 폴백:
                # /?mo=XXX (cl 없음) URL 이 sorisem 측에서 카테고리 콘텐츠 없이
                # 메인 사이드바만 응답하는 경우(자료실·전자도서관 등), URL 을
                # /?mo=XXX&cl=XXX 형태로 한 번 더 시도해 본다. cl 패턴은 7·8번
                # 카테고리에서 sorisem 이 사용하는 형식과 동일.
                if self._looks_like_main_sidebar_only(merged) and "cl=" not in board_url:
                    m_mo = re.search(r"[?&]mo=([a-zA-Z0-9_]+)", board_url)
                    if m_mo:
                        mo_val = m_mo.group(1)
                        retry_url = (
                            board_url + ("&" if "?" in board_url else "?")
                            + f"cl={mo_val}"
                        )
                        try:
                            full_retry = (
                                retry_url if retry_url.startswith("http")
                                else f"{SORISEM_BASE_URL}{retry_url}"
                            )
                            resp = self.session.get(full_retry, timeout=15)
                            html2 = resp.text
                        except Exception:
                            html2 = ""
                        if html2 and len(html2) > 50:
                            sm2 = parse_sub_menus(html2, base_url=retry_url)
                            fb2 = self._extract_fallback_links(html2)
                            merged2: list = []
                            seen2: set[str] = set()
                            for it2 in list(sm2 or []) + list(fb2 or []):
                                if it2.url in seen2:
                                    continue
                                seen2.add(it2.url)
                                merged2.append(it2)
                            if merged2 and not self._looks_like_main_sidebar_only(merged2):
                                board_url = retry_url
                                merged = merged2
                                html = html2

                # 디버그 덤프 — bo_table 이 없는 모든 페이지를 data/ 로 저장.
                try:
                    from config import DATA_DIR
                    safe = re.sub(r"[^A-Za-z0-9]+", "_", board_url)[:40]
                    os.makedirs(DATA_DIR, exist_ok=True)
                    with open(
                        os.path.join(DATA_DIR, f"category_{safe}.html"),
                        "w", encoding="utf-8",
                    ) as _df:
                        _df.write(html)
                except Exception:
                    pass

                if merged:
                    self._show_sub_menu(merged, name, base_url=board_url)
                    return
                speak(f"{name}에 표시할 내용이 없습니다.")
                self.status_bar.SetStatusText("준비", 0)
                return

            # 1순위: 게시글 목록
            posts = parse_board_list(html)
            if posts:
                self.current_board_url = board_url
                self._show_post_list(posts, name, board_url, 1)
                return

            # v1.7 — 자동 폴백:
            # (a) "게시판 접근권한이 없습니다" 응답이면 hub URL (/?mo=XXX&cl=XXX)
            #     을 한 번 GET 해 세션 컨텍스트를 만들고 원본을 재요청.
            # (b) 여전히 글이 0건이면 cl= 를 떼고 다시 시도.
            access_denied = (
                "접근권한이 없습니다" in html
                or "오류안내" in html
                or "history.back()" in html
            )
            if (
                "bo_table=" in board_url
                and re.search(r"[?&]cl=([^&#]+)", board_url)
                and (not posts or access_denied)
            ):
                m_cl = re.search(r"[?&]cl=([^&#]+)", board_url)
                cl_val = m_cl.group(1) if m_cl else ""

                # (a) hub 워밍업 후 원본 재시도
                if cl_val and access_denied:
                    try:
                        hub_url = f"{SORISEM_BASE_URL}/?mo={cl_val}&cl={cl_val}"
                        self.session.get(hub_url, timeout=15)
                        full_orig = (
                            board_url if board_url.startswith("http")
                            else f"{SORISEM_BASE_URL}{board_url}"
                        )
                        resp = self.session.get(full_orig, timeout=15)
                        html_retry = resp.text
                    except Exception:
                        html_retry = ""
                    if html_retry and len(html_retry) > 50:
                        posts_retry = parse_board_list(html_retry)
                        if posts_retry:
                            self.current_board_url = board_url
                            self._show_post_list(posts_retry, name, board_url, 1)
                            return
                        # 재시도 결과로 html 갱신해 (b) 폴백에 사용
                        html = html_retry

                # (b) cl= 파라미터 제거 후 재시도
                retry_url = re.sub(r"[?&]cl=[^&#]+", "", board_url)
                retry_url = retry_url.replace("?&", "?").rstrip("?&")
                try:
                    full_retry = (
                        retry_url if retry_url.startswith("http")
                        else f"{SORISEM_BASE_URL}{retry_url}"
                    )
                    resp = self.session.get(full_retry, timeout=15)
                    html2 = resp.text
                except Exception:
                    html2 = ""
                if html2 and len(html2) > 50:
                    posts2 = parse_board_list(html2)
                    if posts2:
                        self.current_board_url = board_url
                        self._show_post_list(posts2, name, board_url, 1)
                        return

            # URL에 bo_table이 있으면 게시판 → 글이 0개인 빈 게시판
            if "bo_table=" in board_url:
                # 진단용: 빈 게시판으로 판정된 HTML 을 data/ 로 덤프해
                # parse 실패인지 진짜 0건인지 사용자가 공유할 수 있게 한다.
                try:
                    from config import DATA_DIR
                    safe = re.sub(r"[^A-Za-z0-9]+", "_", board_url)[:40]
                    os.makedirs(DATA_DIR, exist_ok=True)
                    with open(
                        os.path.join(DATA_DIR, f"empty_board_{safe}.html"),
                        "w", encoding="utf-8",
                    ) as _df:
                        _df.write(html)
                except Exception:
                    pass

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
            sub_menus = parse_sub_menus(html, base_url=board_url)
            if sub_menus:
                self._show_sub_menu(sub_menus, name, base_url=board_url)
                return

            # 3순위: 본문
            content = parse_post_content(html)
            if content and content.body:
                self._show_post_dialog(content)
                self.status_bar.SetStatusText("준비", 0)
                return

            # 4순위: 페이지의 모든 의미있는 링크를 하위메뉴로 표시
            fallback_menus = self._extract_fallback_links(html)
            if fallback_menus:
                self._show_sub_menu(fallback_menus, name, base_url=board_url)
                return

            speak(f"{name}에 표시할 내용이 없습니다.")
            self.status_bar.SetStatusText("준비", 0)

        self._fetch_page(url, on_loaded)

    def _looks_like_main_sidebar_only(self, items: list) -> bool:
        """페이지가 메인 메뉴 사이드바만 노출했는지 휴리스틱 판정.

        items 의 절반 이상이 `menu_manager.menus` 의 URL 이면 카테고리 콘텐츠
        없이 사이드바만 돌려준 페이지로 본다. /?mo=pds, /?mo=lib2013 같은 빈
        카테고리 응답에서 자동 폴백을 트리거하는 데 사용.
        """
        if not items:
            return True
        try:
            main_urls = {m.url for m in self.menu_manager.menus}
        except Exception:
            return False
        if not main_urls:
            return False
        n_total = 0
        n_main = 0
        for it in items:
            url = (it.url or "").strip()
            if not url or url in ("/", "", "#"):
                continue
            if "wr_id=" in url.lower():
                continue
            n_total += 1
            # 절대 URL 도 상대로 정규화해 비교
            check_url = url
            if check_url.startswith("http") and SORISEM_BASE_URL in check_url:
                check_url = check_url.replace(SORISEM_BASE_URL, "")
            if check_url in main_urls:
                n_main += 1
        if n_total == 0:
            return True
        # 절반 이상이 메인 메뉴 URL이면 사이드바만 응답한 것으로 판정
        return n_main >= max(2, n_total // 2)

    def _extract_fallback_links(self, html: str) -> list:
        """페이지에서 노이즈를 제외한 의미 있는 모든 `<a>` 링크를 SubMenuItem
        리스트로 반환. parse_sub_menus 의 셀렉터로 잡히지 않는 페이지(예:
        /?mo=potion 같은 카테고리 랜딩) 에서 폴백으로 사용한다.
        """
        from bs4 import BeautifulSoup as _BS
        from page_parser import SubMenuItem
        _soup = _BS(html, "html.parser")
        for tag in _soup.find_all(["script", "style", "footer"]):
            tag.decompose()

        noise_texts = [
            "본문으로", "상단으로", "로그아웃", "개인정보", "이용약관",
            "돌아가기", "메일", "쪽지", "검색", "홈", "상단", "맨위",
            "저작권", "copyright", "top", "skip", "동사무소",
        ]
        noise_hrefs = [
            "login", "logout", "register", "memo.php", "formmail",
            "mailto:", "password", "javascript:", "history.back",
        ]
        out = []
        seen = set()
        for a in _soup.find_all("a", href=True):
            href = a.get("href", "").strip()
            text = a.get_text(strip=True)
            if not text or len(text) < 2 or len(text) > 60:
                continue
            if href in ("#", ""):
                continue
            href_lower = href.lower()
            # 게시글 본문 링크는 하위 메뉴가 아니라 게시글 자체 — 제외.
            # /?mo=potion 같은 카테고리 페이지가 보드로 리다이렉트되었을 때
            # 게시글들이 하위 메뉴로 잘못 표시되는 문제 차단.
            if "wr_id=" in href_lower:
                continue
            if any(k in href_lower for k in noise_hrefs):
                continue
            if any(k in text for k in noise_texts):
                continue
            if href.startswith("http") and SORISEM_BASE_URL not in href:
                pass
            elif href.startswith("http"):
                href = href.replace(SORISEM_BASE_URL, "")
            if href not in seen:
                seen.add(href)
                out.append(SubMenuItem(text, href))
        return out

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

        # Ctrl+B: 즐겨찾기 열기 — v1.7
        elif keycode == ord("B") and ctrl and not shift:
            self.on_open_bookmarks()
            return

        # Ctrl+D: 현재 게시판/게시물 즐겨찾기에 추가 — v1.7
        elif keycode == ord("D") and ctrl and not shift and not alt:
            self.on_add_bookmark()
            return

        # Ctrl+P: 명령 도구 모음 — v1.7
        elif keycode == ord("P") and ctrl and not shift and not alt:
            self.on_open_command_palette()
            return

        # Ctrl+Shift+S: 현재 게시판 구독 토글 — v1.7
        elif keycode == ord("S") and ctrl and shift and not alt:
            self.on_toggle_subscription()
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
        elif keycode in (ord("D"), ord("d")) and alt and not ctrl:
            self._delete_post()
        elif keycode == wx.WXK_DELETE:
            if self.current_view == VIEW_POST_LIST:
                self._delete_post()

        # D (단독): 게시물 목록에서 첨부파일 자동 다운로드 — v1.7
        elif (
            keycode in (ord("D"), ord("d"))
            and not ctrl and not alt and not shift
        ):
            if self.current_view == VIEW_POST_LIST:
                self._download_post_attachments_from_list()
            else:
                event.Skip()

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
        dialog = WriteDialog(
            self, self.session, bo_table,
            user_rank=self.current_user_rank,
        )
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

    def _on_menu_nas_logout(self, event):
        """메뉴바 '도구 > 초록등대 자료실 로그아웃' 핸들러.

        확인 대화상자를 거친 뒤 마운트된 드라이브를 분리하고 저장된 NAS
        자격증명(아이디·비밀번호)을 삭제한다. 다음 연결 시도 시 사용자가 다시
        입력해야 한다.
        """
        try:
            from nas import (
                get_mounted_drive, unmount,
                load_nas_credentials, delete_nas_credentials,
            )
        except Exception as e:
            speak(f"NAS 모듈을 불러올 수 없습니다. {e}")
            return

        has_creds = bool(load_nas_credentials())
        drive = get_mounted_drive()
        if not has_creds and not drive:
            speak("저장된 자료실 자격증명이 없습니다.")
            wx.MessageBox(
                "저장된 자료실 자격증명이 없습니다.\n로그아웃할 정보가 없습니다.",
                "안내", wx.OK | wx.ICON_INFORMATION, self,
            )
            return

        ans = wx.MessageBox(
            "초록등대 자료실에서 로그아웃할까요?\n\n"
            "마운트된 드라이브를 분리하고 저장된 자료실 아이디·비밀번호를\n"
            "삭제합니다. 다음에 연결할 때 다시 입력해야 합니다.",
            "초록등대 자료실 로그아웃",
            wx.YES_NO | wx.ICON_QUESTION, self,
        )
        if ans != wx.YES:
            return

        try:
            if drive:
                unmount(drive)
        except Exception:
            pass
        try:
            delete_nas_credentials()
        except Exception:
            pass

        speak("초록등대 자료실에서 로그아웃했습니다.")
        wx.MessageBox(
            "초록등대 자료실에서 로그아웃했습니다.",
            "완료", wx.OK | wx.ICON_INFORMATION, self,
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
            "base_url": getattr(self, "current_sub_menu_url", "") or "",
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

    def _download_post_attachments_from_list(self):
        """게시물 목록에서 D 키로 호출. 현재 줄의 게시물 첨부파일을
        모두 다운로드 폴더에 저장한다. 게시물 본문 창을 띄우지 않음.
        """
        if self.current_view != VIEW_POST_LIST:
            return
        index = self._get_current_line_index()
        if index < 0 or index >= len(self.current_posts):
            speak("선택된 게시물이 없습니다.")
            return
        post = self.current_posts[index]

        speak(f"{post.title} 첨부파일을 확인하는 중입니다.")
        self.status_bar.SetStatusText("첨부파일 확인 중...", 0)

        def on_loaded(html, error):
            self.status_bar.SetStatusText("준비", 0)
            if error or not html:
                speak("게시물을 불러올 수 없습니다.")
                return
            content = parse_post_content(html)
            if not content:
                speak("게시물 내용을 불러올 수 없습니다.")
                return
            files = getattr(content, "files", None) or []
            if not files:
                speak("이 게시물에는 첨부파일이 없습니다.")
                wx.MessageBox(
                    "이 게시물에는 첨부파일이 없습니다.",
                    "안내", wx.OK | wx.ICON_INFORMATION, self,
                )
                return
            self._download_files_to_folder(content, files)

        self._fetch_page(post.url, on_loaded)

    def _download_files_to_folder(self, content, files: list):
        """주어진 첨부파일 목록을 백그라운드로 다운로드 폴더에 저장.

        PostDialog.on_download_file 와 동일한 로직 — DAISY 자동 변환 안내까지
        포함. 게시물 목록 D 단축키와 명령 도구 모음 등에서 공유.
        """
        from config import get_download_dir
        try:
            from sound import play_event
        except Exception:
            play_event = None

        download_dir = get_download_dir()
        total = len(files)
        speak(f"첨부파일 다운로드를 시작합니다. {total}개 파일")
        if play_event:
            try:
                play_event("download_start")
            except Exception:
                pass

        from post_dialog import PostDialog as _PD

        def _beep(freq):
            try:
                import winsound
                winsound.Beep(freq, 100)
            except Exception:
                pass

        def worker():
            success = 0
            fail = 0
            saved_paths: list[str] = []
            for fi in files:
                url = fi["url"]
                raw_name = fi["name"]
                clean_name = _PD._clean_filename(raw_name)
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
                            if total_size > 0:
                                pct = int(downloaded / total_size * 100)
                                if pct >= last_pct + 10:
                                    last_pct = (pct // 10) * 10
                                    freq = 400 + last_pct * 6
                                    _beep(freq)
                    dl_entry["status"] = "완료"
                    _beep(1200)
                    success += 1
                    saved_paths.append(save_path)
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

            # DAISY ZIP 자동 변환 안내 — PostDialog 와 동일한 흐름 재사용.
            for path in saved_paths:
                try:
                    from daisy import is_daisy_zip
                    if is_daisy_zip(path):
                        wx.CallAfter(self._offer_daisy_for_path, path)
                        break
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _offer_daisy_for_path(self, zip_path: str):
        """DAISY ZIP 자동 변환 안내. 사용자가 수락하면 _convert_daisy_zip 호출."""
        try:
            from daisy import convert_zip_to_text
        except Exception:
            return
        ans = wx.MessageBox(
            "DAISY 도서로 보이는 ZIP 파일을 다운로드했습니다.\n"
            "본문 텍스트로 변환할까요?",
            "DAISY 도서 변환", wx.YES_NO | wx.ICON_QUESTION, self,
        )
        if ans != wx.YES:
            return
        try:
            out_dir = convert_zip_to_text(zip_path)
            speak("DAISY 도서 변환이 완료되었습니다.")
            wx.MessageBox(
                f"변환 결과 폴더:\n{out_dir}",
                "DAISY 변환 완료", wx.OK | wx.ICON_INFORMATION, self,
            )
        except Exception as e:
            speak("DAISY 변환에 실패했습니다.")
            wx.MessageBox(
                f"DAISY 변환에 실패했습니다.\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

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
            self._show_sub_menu(
                prev["sub_menus"], prev["menu_name"],
                base_url=prev.get("base_url", ""),
            )
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

    # ── 게시물 작성자 검증 ──

    def _names_match(self, a: str | None, b: str | None) -> bool:
        """닉네임 두 개가 같은 사람을 가리키는지 판단.
        공백·대소문자·말미 존칭(님/씨) 을 정규화한 뒤 비교하고, 한쪽이 다른
        쪽을 포함하는 경우("닉네임" ⊂ "닉네임 (아이디)") 도 일치로 본다."""
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

    def _verify_post_ownership(self, bo_table: str, wr_id: str,
                               hint_author_name: str = "") -> tuple[bool, str]:
        """게시물 작성자와 현재 로그인한 사용자가 동일한지 확인한다.

        서버는 동호회 관리자에게 모든 게시물의 수정/삭제 권한을 주지만, 이 앱은
        "자기 글만 수정·삭제한다" 정책을 클라이언트에서 강제하기 위해 아래 순서로
        판단한다:

          1) 닉네임 비교 (가장 빠르고 안정적)
             - self.current_user_nickname vs hint_author_name(게시물 목록의 닉네임)
             - 일치 → 본인 / 불일치 → 타인
          2) 게시물 본문 페이지에서 작성자 mb_id 추출 후 current_user_id 와 비교
          3) 둘 다 판단 불가면 허용으로 폴백 (본인 글 막히는 것 방지)
             — 이 경로는 서버의 기본 권한 체크가 계속 동작하므로 관리자 이외엔
             의미가 없고, 관리자 계정에서만 드물게 false-negative 발생.

        Returns:
            (True,  ""):  본인 게시물 → 진행 허용
            (False, msg): 타인 게시물 → 진행 거부 (msg 는 안내용)
        """
        if not self.current_user_id:
            return False, "로그인 사용자 정보를 확인할 수 없어 수정·삭제를 중단합니다."

        # 1) 닉네임 비교 (HTTP 호출 없음)
        if self.current_user_nickname and hint_author_name:
            if self._names_match(self.current_user_nickname, hint_author_name):
                return True, ""
            return False, "본인이 작성한 게시물만 수정·삭제할 수 있습니다."

        # 2) 게시물 본문 페이지에서 mb_id 추출
        try:
            post_url = (
                f"{SORISEM_BASE_URL}/bbs/board.php?"
                f"bo_table={bo_table}&wr_id={wr_id}"
            )
            resp = self.session.get(post_url, timeout=15)
        except Exception as e:
            # 네트워크 에러: 진행 허용 (본인 글 막힘 방지). 서버가 최종 권한 검사.
            return True, ""

        from page_parser import extract_post_author_id
        author_id = extract_post_author_id(resp.text)

        if author_id:
            if author_id.strip().lower() == self.current_user_id.strip().lower():
                return True, ""
            return False, "본인이 작성한 게시물만 수정·삭제할 수 있습니다."

        # 3) 판단 불가 — 서버 기본 권한 체크에 맡긴다.
        return True, ""

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

        # 클라이언트측 작성자 검증: 서버가 관리자에게 허용하더라도 본인 글만
        # 수정할 수 있게 일관된 정책을 강제.
        owned, own_msg = self._verify_post_ownership(
            bo_table, wr_id, hint_author_name=getattr(post, "author", ""),
        )
        if not owned:
            speak(own_msg)
            wx.MessageBox(own_msg, "수정 불가", wx.OK | wx.ICON_WARNING, self)
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
                user_rank=self.current_user_rank,
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

        # 클라이언트측 작성자 검증 — 확인 대화상자를 띄우기 전에 먼저 검사해야
        # "확인 → 본인 글 아님" 순서의 어색한 흐름을 피할 수 있다.
        owned, own_msg = self._verify_post_ownership(
            bo_table, wr_id, hint_author_name=getattr(post, "author", ""),
        )
        if not owned:
            speak(own_msg)
            wx.MessageBox(own_msg, "삭제 불가", wx.OK | wx.ICON_WARNING, self)
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
            # 게시판·클럽 페이지의 액션 버튼 텍스트 — 페이지 제목으로 부적합
            "글쓰기", "쓰기", "답변", "답글", "수정", "삭제",
            "이전", "다음", "처음", "마지막",
            "더보기", "전체", "보기", "닫기", "취소",
            "확인", "저장", "전송", "신청", "등록",
            "위로", "아래로", "이동",
            "댓글", "댓글쓰기", "추천", "비추천",
            "스크랩", "공유", "프린트", "인쇄",
            "관리자", "운영자",
            # _show_sub_menu 의 header_noise 와 동기화
            "게시판 관리", "게시판관리", "멀티업로드",
            "img", "철머", "로그아웃", "돌아가기",
            "소리샘 동사무소", "동사무소",
            "회원가입", "마이페이지", "쪽지", "쪽지함",
            "메일", "메일함", "알림", "구독",
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
            # "N. X" / "N) X" / "N X"(예: "2 페이지") 형식은 하위 메뉴 항목·
            # 페이지 번호 링크 등 — 클럽/게시판 제목으로 부적합.
            if re.match(r'^\d+\s*([\.\)]|\s|페이지|page|$)', t, re.IGNORECASE):
                return False
            return True

        # ⭐ 최우선: 현재 코드를 가리키는 링크의 텍스트
        # 예: <a href="/plugin/ar.club/?cl=hims">셀바스헬스케어(구) 힘스인터네셔널</a>
        # 주의: cl=hims 가 포함된 하위 게시판 URL(bo_table=xxx&cl=hims)은 제외
        if match_code:
            # 우선순위 별 후보 보관 — strict 가 우선.
            strict: list[str] = []      # /plugin/ar.club/?cl=CODE 정확 매칭
            board_strict: list[str] = []  # /bbs/board.php?bo_table=CODE (글쓰기·답글 등 wr_id 제외)
            loose: list[str] = []       # 그 외 cl=CODE 포함 링크 (mo=… 카테고리)
            for a in soup.find_all("a", href=True):
                href = a.get("href", "")
                href_lower = href.lower()

                # wr_id 가 있으면 게시물 본문 링크 — 제목으로 부적합
                if "wr_id=" in href_lower:
                    continue

                t = a.get_text(" ", strip=True)
                if not t:
                    t = a.get("title", "").strip()
                if not _is_valid(t):
                    continue

                # 페이지 번호 링크(`page=N`) 제외 — 텍스트가 "2", "3" 또는
                # "2 페이지" 등 형식이라 클럽 제목으로 부적합.
                if re.search(r"[?&]page=\d", href_lower):
                    continue

                # 1) 클럽 메인 (가장 신뢰: ar.club 플러그인 진입점)
                if (
                    "/plugin/ar.club/" in href_lower
                    and f"cl={match_code}" in href_lower
                    and "bo_table=" not in href_lower
                ):
                    strict.append(t)
                    continue

                # 2) 게시판 메인 — bo_table=CODE 가 정확히 들어 있고 wr_id 없음
                if f"bo_table={match_code}" in href_lower:
                    board_strict.append(t)
                    continue

                # 3) 카테고리 nav (mo=…&cl=CODE) — 본 페이지 내부의 sub-link 일 수
                #    있어 부정확한 후보. 폴백으로만 사용.
                if (
                    f"cl={match_code}" in href_lower
                    and "bo_table=" not in href_lower
                ):
                    loose.append(t)

            # 우선순위에 따라 첫 후보를 채택. 길이 휴리스틱은 sub-link 가 더 긴
            # 이름인 경우 클럽 본명을 가리는 부작용이 있어 "맨 처음" 으로 변경.
            for bucket in (strict, board_strict, loose):
                if bucket:
                    return bucket[0]

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
        """사용자가 바로가기 대화상자에 입력한 코드로 이동.

        v1.7 — 메뉴 클릭 흐름과 동일한 동작 보장:
        sorisem 사이트 구조상 hims/xvsrd 같은 클럽은 단일 부모 hub 가 아니라
        여러 카테고리 hub (`/?mo=potion`, `/?mo=prg`, `/plugin/ar.club/?cl=green`
        등) 중 하나의 자식으로 등록되어 있다. 여러 후보 hub 의 HTML 을 순차로
        가져와 각 `<a href>` 의 바로가기 코드가 입력 코드와 일치하면 그
        sub-menu 의 실제 URL/이름으로 `_load_and_show` 를 호출한다.

        모든 후보 hub 에서 못 찾으면 `/bbs/board.php?bo_table=CODE` 직접 호출
        폴백. 그래도 없으면 "표시할 내용이 없습니다" 안내.
        """
        board_url = f"/bbs/board.php?bo_table={code}"
        self.status_bar.SetStatusText(f"{code} 검색 중...", 0)
        speak(f"{code} 검색 중입니다.")

        code_lower = code.lower()

        def resolve_display_name(html) -> str:
            display_name = self._extract_page_title(html, "", match_code=code)
            return display_name or code

        def render_board(html, tried_url) -> bool:
            if not html or len(html) < 100:
                return False
            if (
                'alert("게시판 접근권한이 없습니다' in html
                or "<title>오류안내" in html
            ):
                return False
            posts = parse_board_list(html)
            sub_menus = parse_sub_menus(html, base_url=tried_url)
            display_name = resolve_display_name(html)
            if posts:
                self.current_board_url = tried_url
                self._show_post_list(posts, display_name, tried_url, 1)
                return True
            if sub_menus:
                self._show_sub_menu(sub_menus, display_name, base_url=tried_url)
                return True
            if "bo_table=" in tried_url:
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
            if error or not render_board(html, board_url):
                speak("표시할 내용이 없습니다.")
                self.status_bar.SetStatusText("준비", 0)

        # 후보 부모 hub 목록 — sorisem 의 주요 카테고리/클럽 hub.
        # 일반 클럽(hims, xvsrd 등) 은 `/?mo=potion` (동호회) 아래에 가장
        # 흔히 등록되어 있어 후순위에 둔다. green 은 가장 빈번한 케이스라 우선.
        candidate_hubs = [
            "/plugin/ar.club/?cl=green",
            "/?mo=potion&cl=potion",
            "/?mo=prg&cl=prg",
            "/?mo=pds&cl=pds",
            "/?mo=magazin&cl=magazin",
            "/?mo=lib2013&cl=lib2013",
            "/?mo=edu2013&cl=edu2013",
            "/?mo=braille",
        ]

        def search_in_hubs():
            """각 후보 hub 의 HTML 을 순차 GET → 코드 매칭 검색.

            발견 시 (href, text) 튜플 반환. 못 찾으면 (None, None).
            진단용으로 모든 hub 응답을 `data/goto_hub_<safe>.html` 로 dump.
            """
            from bs4 import BeautifulSoup as _BS
            from menu_manager import extract_shortcut_code
            try:
                from config import DATA_DIR
                import os as _os
                _os.makedirs(DATA_DIR, exist_ok=True)
            except Exception:
                DATA_DIR = None

            target_patterns = (
                f"cl={code_lower}",
                f"bo_table={code_lower}",
                f"mo={code_lower}",
            )

            for hub in candidate_hubs:
                try:
                    full = (
                        hub if hub.startswith("http")
                        else f"{SORISEM_BASE_URL}{hub}"
                    )
                    resp = self.session.get(full, timeout=15)
                    html = resp.text
                except Exception:
                    continue

                # 진단 dump
                if DATA_DIR:
                    try:
                        safe = re.sub(r"[^A-Za-z0-9]+", "_", hub)[:50]
                        with open(
                            os.path.join(DATA_DIR, f"goto_hub_{safe}.html"),
                            "w", encoding="utf-8",
                        ) as _f:
                            _f.write(html)
                    except Exception:
                        pass

                if not html or len(html) < 100:
                    continue

                try:
                    soup = _BS(html, "html.parser")
                except Exception:
                    continue

                for a in soup.find_all("a", href=True):
                    href = a.get("href", "")
                    href_low = href.lower()
                    if not href or "wr_id=" in href_low:
                        continue

                    # 1) extract_shortcut_code 로 정확 매칭
                    sm_code = (extract_shortcut_code(href) or "").lower()
                    matched = (sm_code == code_lower)

                    # 2) URL 안에 코드가 있는 query param 으로 매칭 (보조)
                    if not matched:
                        matched = any(p in href_low for p in target_patterns)

                    if matched:
                        text = a.get_text(" ", strip=True)
                        if not text:
                            text = a.get("title", "").strip()
                        if not text:
                            continue
                        return href, text
            return None, None

        def worker():
            try:
                match_url, match_name = search_in_hubs()
            except Exception:
                match_url, match_name = None, None

            if match_url:
                wx.CallAfter(self.status_bar.SetStatusText, "준비", 0)
                wx.CallAfter(
                    self._load_and_show, match_url, match_name or code,
                )
                return

            # 후보 hub 에서 매칭 실패 — 게시판 URL 직접 시도.
            wx.CallAfter(self._fetch_page, board_url, on_board_loaded)

        threading.Thread(target=worker, daemon=True).start()

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

    # ── 사용자 편집 메뉴 파일 ──

    def on_edit_menu_file(self, event):
        """사용자 편집용 메뉴 텍스트 파일을 기본 편집기로 연다.

        파일이 없으면 현재 메뉴 목록을 시드로 기록한 뒤 연다.
        """
        from config import MENU_LIST_TXT_FILE
        try:
            self.menu_manager.export_to_txt()
        except Exception as e:
            speak(f"메뉴 파일을 준비하지 못했습니다. {e}")
            return

        if not os.path.exists(MENU_LIST_TXT_FILE):
            speak("메뉴 파일을 찾을 수 없습니다.")
            return

        try:
            os.startfile(MENU_LIST_TXT_FILE)
            speak("메뉴 목록 파일을 엽니다. 저장 후 다시 읽기 메뉴를 실행하세요.")
        except OSError as e:
            speak(f"파일을 열 수 없습니다. {e}")
            wx.MessageBox(
                f"파일을 열 수 없습니다.\n경로: {MENU_LIST_TXT_FILE}\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

    def on_edit_reply_templates(self, event=None):
        """v1.7 — 답장 템플릿 파일을 기본 편집기로 연다.

        파일이 없으면 기본 템플릿으로 자동 생성한 뒤 연다. 사용자가 저장하면
        다음에 댓글 입력창에서 Alt+1~9 가 새 내용을 사용한다 (별도 reload 불필요).
        """
        from config import REPLY_TEMPLATES_FILE
        from reply_templates import load_templates  # 자동 생성 트리거
        try:
            load_templates()  # 파일이 없으면 기본값으로 생성
        except Exception:
            pass
        if not os.path.exists(REPLY_TEMPLATES_FILE):
            speak("답장 템플릿 파일을 찾을 수 없습니다.")
            return
        try:
            os.startfile(REPLY_TEMPLATES_FILE)
            speak("답장 템플릿 파일을 엽니다. 한 줄에 하나씩 적고 저장하세요.")
        except OSError as e:
            speak(f"파일을 열 수 없습니다. {e}")
            wx.MessageBox(
                f"파일을 열 수 없습니다.\n경로: {REPLY_TEMPLATES_FILE}\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

    def on_reload_menu_file(self, event):
        """메뉴 파일을 다시 읽어 메인 메뉴에 반영."""
        try:
            self.menu_manager.load()
            self._show_main_menu()
            speak(f"메뉴를 다시 불러왔습니다. {len(self.menu_manager.menus)}개.")
        except Exception as e:
            speak(f"메뉴를 다시 불러오지 못했습니다. {e}")
            wx.MessageBox(
                f"메뉴를 다시 불러오지 못했습니다.\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

    def on_reset_menu_file(self, event):
        """사용자 편집 파일을 삭제하고 다음 실행부터 자동 감지 복원.

        현재 세션은 기존 JSON 캐시를 그대로 사용한다.
        """
        from config import MENU_LIST_TXT_FILE
        if not os.path.exists(MENU_LIST_TXT_FILE):
            speak("사용자 메뉴 파일이 없습니다. 이미 자동 감지를 사용하고 있습니다.")
            return

        dlg = wx.MessageDialog(
            self,
            "사용자 메뉴 목록 파일을 삭제합니다.\n"
            "다음 로그인부터는 소리샘 메인 페이지에서 메뉴를 자동으로 다시 가져옵니다.\n\n"
            "계속할까요?",
            "메뉴 목록 초기화",
            wx.YES_NO | wx.ICON_QUESTION,
        )
        confirm = dlg.ShowModal()
        dlg.Destroy()
        if confirm != wx.ID_YES:
            return
        try:
            os.remove(MENU_LIST_TXT_FILE)
            speak("사용자 메뉴 파일을 삭제했습니다. 다음 실행부터 자동 감지됩니다.")
        except OSError as e:
            speak(f"파일을 삭제하지 못했습니다. {e}")
            wx.MessageBox(
                f"파일을 삭제하지 못했습니다.\n{e}",
                "오류", wx.OK | wx.ICON_ERROR, self,
            )

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
        또한 사용자가 명시적으로 확인을 누른 경우이므로 10분 TTL 캐시를
        무시하고 GitHub API 를 강제로 한 번 더 조회한다 — 갓 나온 릴리스를
        놓치지 않기 위함.
        manual=False 이면 새 버전 있을 때만 알림 + 캐시 사용.
        """
        channel = load_update_settings().get("channel", "stable")

        def worker():
            info = check_latest_release(
                channel=channel, use_cache=not manual,
            )
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
                speak(
                    f"현재 사용 중인 버전 {APP_VERSION} 이 최신입니다."
                )
                wx.MessageBox(
                    f"현재 사용 중인 버전이 최신입니다.\n"
                    f"설치 버전 {APP_VERSION}\n"
                    f"최신 버전 {info.version}",
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
        speak(
            f"초록멀티 버전 {info.version} 업데이트가 있습니다. "
            f"현재 버전은 {APP_VERSION}입니다. 업데이트하시겠습니까?"
        )
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

    @staticmethod
    def _update_beep(percent: int):
        """업데이트 진행률이 10% 단위로 넘어갈 때 짧은 비프 재생.

        진행률이 높아질수록 주파수가 올라가 청각적으로 진행을 인지하도록 함.
        사운드 마스터 스위치가 꺼져 있으면 재생 안 함.
        """
        try:
            from sound import load_sound_settings
            if not load_sound_settings().get("enabled", True):
                return
        except Exception:
            pass
        try:
            import winsound
            # 0 → 500Hz, 100 → 1500Hz
            freq = 500 + max(0, min(100, percent)) * 10
            winsound.Beep(freq, 60)
        except Exception:
            pass

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

        state = {"cancelled": False, "error": None, "path": None, "bucket": -1}

        def progress_cb(downloaded: int, total: int) -> bool:
            # 워커 스레드에서 호출됨. UI 갱신은 CallAfter.
            # 10% 단위(bucket)로 넘어갈 때마다 비프 재생.
            if total > 0:
                percent = int(downloaded / total * 100)
                bucket = percent // 10
                if bucket > state["bucket"]:
                    state["bucket"] = bucket
                    wx.CallAfter(self._update_beep, bucket * 10)
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
        state = {"cancelled": False, "error": None, "path": None, "bucket": -1}

        def progress_cb(downloaded, total):
            if total > 0:
                percent = int(downloaded / total * 100)
                bucket = percent // 10
                if bucket > state["bucket"]:
                    state["bucket"] = bucket
                    wx.CallAfter(self._update_beep, bucket * 10)
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
        extract_state = {"cancelled": False, "error": None, "bucket": -1}

        def xprog(done, total):
            if total > 0:
                percent = int(done / total * 100)
                bucket = percent // 10
                if bucket > extract_state["bucket"]:
                    extract_state["bucket"] = bucket
                    wx.CallAfter(self._update_beep, bucket * 10)
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
        state = {"cancelled": False, "error": None, "path": None, "bucket": -1}

        def progress_cb(downloaded, total):
            if total > 0:
                percent = int(downloaded / total * 100)
                bucket = percent // 10
                if bucket > state["bucket"]:
                    state["bucket"] = bucket
                    wx.CallAfter(self._update_beep, bucket * 10)
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
            "  · 메일함·쪽지함 목록에서도 동일 동작 (상태/보낸이/날짜/제목)",
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
            "D: 첨부파일 즉시 다운로드 (게시물 목록에서, 본문 창을 열지 않음)  — v1.7",
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
            "=== 팝업(컨텍스트) 메뉴 — v1.6 추가 ===",
            "메뉴 키 / Shift+F10 / 마우스 우클릭으로 상황별 메뉴 호출",
            "  · 게시물 목록: 작성·수정·삭제·검색·새로고침",
            "  · 게시물 본문: 본문 저장·첨부 저장·URL 목록·댓글 작성·",
            "                 게시물 수정/삭제/답변·이전/다음 게시물",
            "  · 댓글 목록: 댓글 작성·수정·삭제·정렬 순서 변경",
            "  · 메일함/쪽지함: 새 글 쓰기·답장·읽기·삭제·새로고침",
            "  · 메일/쪽지 보기: 답장·삭제·첨부 저장·본문 저장",
            "(이동·페이지 탐색 등 내비게이션은 팝업에 포함하지 않음)",
            "",
            "=== 댓글 입력창 ===",
            "Enter: 줄바꿈 추가",
            "Ctrl+Enter: 댓글 등록 (확인)",
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
            "Shift+좌/우: 필드 순회 (상태·보낸이·날짜·제목)  — v1.6",
            "Enter: 쪽지/메일 열기",
            "D 또는 Delete: 선택 항목 삭제",
            "Shift+Delete: 현재 함 전체 비우기",
            "R: 답장 (받은함)",
            "N: 새 쪽지/메일 작성",
            "F: 새로고침",
            "PageDown/PageUp: 다음 페이지 누적 로드",
            "Alt+R / Alt+S: 받은함 / 보낸함 전환",
            "Alt+A: 모든 쪽지/메일 삭제",
            "(목록 항목 앞에 \"안 읽음 / 읽음\" 상태가 표시됨 — v1.6)",
            "",
            "=== 쪽지·메일 보기 창 ===",
            "PageUp/PageDown, Alt+P/Alt+N: 이전/다음 항목",
            "R: 답장  D/Delete: 삭제  Esc: 닫기",
            "(메일) B: 본문 저장  Alt+S: 첨부 선택 저장  Alt+Shift+S: 모든 첨부 저장",
            "(메일 보기 창 제목에 메일 제목 표시 — v1.6)",
            "",
            "=== 알림 센터 (Ctrl+Shift+N) ===",
            "Enter: 선택 항목 열기",
            "D 또는 Delete: 선택 알림 지우기",
            "A: 모든 알림 지우기",
            "F: 새로고침",
            "(시작 시 자동으로 미확인 메일·쪽지 알림 — v1.6)",
            "",
            "=== 화면 설정 (저시력 지원) ===",
            "F7: 설정 창 열기 (테마·글꼴·사운드 통합)",
            "F6: 다음 테마로 변경",
            "Shift+F6: 이전 테마로 변경",
            "Ctrl++: 글꼴 크게 (확대)",
            "Ctrl+-: 글꼴 작게 (축소)",
            "Ctrl+0: 글꼴 크기 원래대로",
            "",
            "=== v1.7 단축키 ===",
            "Ctrl+B: 즐겨찾기 열기",
            "Ctrl+D: 현재 위치를 즐겨찾기에 추가",
            "Ctrl+P: 명령 도구 모음 (모든 기능을 키워드로 검색·실행)",
            "Ctrl+Shift+S: 현재 게시판 구독 / 구독 해제",
            "Ctrl+Alt+L: 구독 목록 보기",
            "Ctrl+Alt+D: DAISY 도서 변환 대화상자 열기",
            "D: 게시물 목록에서 첨부파일 즉시 다운로드",
            "Alt+1 ~ Alt+9: 댓글 입력창에서 빠른 답장 템플릿 삽입",
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

    # ── v1.7 즐겨찾기 ──

    def on_add_bookmark(self, event=None):
        """Ctrl+D — 현재 위치를 즐겨찾기에 추가.

        화면별 동작:
          · 메인 메뉴: 현재 커서가 있는 메뉴 항목
          · 하위 메뉴: 현재 커서가 있는 하위 메뉴 항목 (상위 메뉴 X)
          · 게시판 글 목록: 현재 게시판 자체
        """
        url = ""
        suggested = ""
        type_hint = "board"

        if self.current_view == VIEW_MAIN_MENU:
            idx = getattr(self, "current_index", 0)
            menu_item = self.menu_manager.get_menu_by_index(idx)
            if menu_item:
                url = menu_item.url
                suggested = re.sub(r"^\s*\d+[\.\)]\s*", "", menu_item.name).strip()
                type_hint = menu_item.type or "board"
            else:
                speak("현재 항목을 알 수 없어 즐겨찾기에 추가하지 못했습니다.")
                return
        elif self.current_view == VIEW_SUB_MENU:
            # current_index 는 textctrl 줄 번호. 0 번 줄은 "메인 메뉴로 돌아가기" 안내,
            # 1 번 줄부터가 실제 하위 메뉴 항목 (current_sub_menus[0]).
            idx = getattr(self, "current_index", 0)
            actual = idx - 1
            sub_items = getattr(self, "current_sub_menus", []) or []
            if 0 <= actual < len(sub_items):
                sub = sub_items[actual]
                url = sub.url if not sub.url.startswith("http") or SORISEM_BASE_URL in sub.url else sub.url
                # 사이트 내부 URL 은 상대 경로로 정규화
                if url.startswith("http") and SORISEM_BASE_URL in url:
                    url = url.replace(SORISEM_BASE_URL, "")
                suggested = re.sub(r"^\s*\d+[\.\)]\s*", "", sub.name).strip()
                type_hint = "board"
            else:
                # 안내 줄이거나 잘못된 인덱스 — 안내
                speak("저장할 하위 메뉴 항목을 선택해 주세요.")
                wx.MessageBox(
                    "즐겨찾기에 추가하려는 하위 메뉴 항목 위에 커서를 두고 "
                    "다시 Ctrl+D 를 눌러 주세요.",
                    "선택 필요",
                    wx.OK | wx.ICON_INFORMATION, self,
                )
                return
        elif self.current_view == VIEW_POST_LIST and self.current_board_url:
            url = self.current_board_url
            suggested = re.sub(
                r"^\s*\d+[\.\)]\s*", "", self.current_menu_name or "게시판",
            ).strip()
            type_hint = "board"
        else:
            speak("현재 화면은 즐겨찾기에 추가할 수 없습니다.")
            wx.MessageBox(
                "지금 화면은 즐겨찾기에 추가할 수 없습니다.\n"
                "메인 메뉴·하위 메뉴 항목 위에서 다시 시도해 주세요.",
                "즐겨찾기 추가 불가",
                wx.OK | wx.ICON_INFORMATION, self,
            )
            return

        if not url:
            speak("주소 정보를 알 수 없어 즐겨찾기에 추가하지 못했습니다.")
            return

        dlg = wx.TextEntryDialog(
            self, "즐겨찾기 이름을 입력하세요.", "즐겨찾기 추가",
            value=suggested,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            name = dlg.GetValue().strip()
        finally:
            dlg.Destroy()
        if not name:
            return

        added = self.bookmark_manager.add(name, url, type_hint)
        if added:
            speak(f"즐겨찾기에 추가했습니다. {name}")
        else:
            speak(f"이미 등록된 즐겨찾기 이름을 갱신했습니다. {name}")

    def on_open_bookmarks(self, event=None):
        """즐겨찾기 대화상자 열기 (Ctrl+B). 선택 시 해당 URL 로 바로 진입."""
        from bookmark_dialog import BookmarkDialog
        dlg = BookmarkDialog(self, self.bookmark_manager)
        code = dlg.ShowModal()
        url = dlg.selected_url
        name = dlg.selected_name
        dlg.Destroy()
        if code == wx.ID_OK and url:
            speak(f"{name} 로 이동합니다.")
            self._load_and_show(url, name)

    # ── v1.7 명령 도구 모음 ──

    def on_open_command_palette(self, event=None):
        """Ctrl+P — 모든 명령을 키워드로 검색·실행."""
        from command_palette import CommandPaletteDialog, Command
        cmds = self._build_command_list()
        dlg = CommandPaletteDialog(self, cmds)
        code = dlg.ShowModal()
        cmd = getattr(dlg, "selected_command", None)
        dlg.Destroy()
        if code == wx.ID_OK and cmd is not None:
            try:
                cmd.callback()
            except Exception as e:
                speak(f"명령 실행 중 오류가 발생했습니다.")
                wx.MessageBox(f"명령 실행 실패: {e}", "오류",
                              wx.OK | wx.ICON_ERROR, self)

    def _build_command_list(self):
        """명령 도구 모음에 표시할 모든 명령 목록 구성."""
        from command_palette import Command
        cmds: list[Command] = []
        cmds.append(Command(
            "메인 메뉴로 이동", "처음 화면으로 돌아갑니다", "Alt+Home",
            lambda: (self._show_main_menu(), speak("메인 메뉴로 돌아왔습니다.")),
        ))
        cmds.append(Command(
            "즐겨찾기 열기", "저장한 즐겨찾기 목록을 봅니다", "Ctrl+B",
            self.on_open_bookmarks,
        ))
        cmds.append(Command(
            "현재 위치를 즐겨찾기에 추가", "현재 화면을 즐겨찾기로 저장", "Ctrl+D",
            self.on_add_bookmark,
        ))
        cmds.append(Command(
            "현재 게시판 구독 토글", "새 글이 올라오면 알림으로 받기", "Ctrl+Shift+S",
            self.on_toggle_subscription,
        ))
        cmds.append(Command(
            "구독 목록 보기", "구독 중인 게시판 관리", "",
            self.on_open_subscriptions,
        ))
        cmds.append(Command(
            "쪽지함 열기", "받은·보낸 쪽지 목록", "Ctrl+M",
            lambda: self.on_open_memo_inbox(None),
        ))
        cmds.append(Command(
            "쪽지 쓰기", "새 쪽지 작성", "Ctrl+Shift+M",
            lambda: self.on_open_memo_compose(None),
        ))
        cmds.append(Command(
            "메일함 열기", "받은·보낸 메일 목록", "Ctrl+Shift+E",
            lambda: self.on_open_mail_compose(None),
        ))
        cmds.append(Command(
            "관리자에게 메일 보내기", "관리자 메일 작성", "Alt+E",
            lambda: self.on_mail(None),
        ))
        cmds.append(Command(
            "알림 센터 열기", "새 메일·쪽지·게시판 글 모음", "Ctrl+Shift+N",
            lambda: self.on_memo_check_now(None),
        ))
        cmds.append(Command(
            "초록등대 자료실 연결", "NAS 마운트", "Ctrl+N",
            lambda: self._on_menu_nas_connect(None),
        ))
        cmds.append(Command(
            "DAISY 도서 변환", "ZIP 으로 받은 DAISY 도서를 압축 해제 + TXT 변환", "",
            self.on_convert_daisy,
        ))
        cmds.append(Command(
            "다운로드 상태", "현재 다운로드 진행 상황", "Ctrl+J",
            lambda: self.on_download_status(None),
        ))
        cmds.append(Command(
            "게시물 검색", "현재 게시판에서 검색", "Ctrl+F",
            self.on_search,
            when=lambda: self.current_view == VIEW_POST_LIST,
        ))
        cmds.append(Command(
            "게시판 새로고침", "현재 게시판 다시 로드", "F5",
            self.on_board_refresh,
            when=lambda: self.current_view == VIEW_POST_LIST,
        ))
        cmds.append(Command(
            "게시물 작성", "새 글 쓰기", "W",
            self._write_post,
            when=lambda: self.current_view == VIEW_POST_LIST,
        ))
        cmds.append(Command(
            "단축키 안내", "전체 단축키 목록", "Ctrl+K",
            lambda: self.on_shortcuts_help(None),
        ))
        cmds.append(Command(
            "사용자 설명서", "전체 설명서 보기", "Shift+F1",
            lambda: self.on_show_manual(None),
        ))
        cmds.append(Command(
            "설정", "테마·글꼴·알림 등 설정", "F7",
            lambda: self.on_show_settings(None),
        ))
        cmds.append(Command(
            "업데이트 확인", "최신 버전 수동 확인", "Alt+U",
            lambda: self.on_manual_update_check(None),
        ))
        cmds.append(Command(
            "로그아웃", "현재 계정 로그아웃", "Ctrl+L",
            lambda: self.on_logout(None),
        ))
        return cmds

    # ── v1.7 게시판 구독 ──

    def on_toggle_subscription(self, event=None):
        """현재 게시판을 구독/구독 해제 (Ctrl+Shift+S).

        구독 중이면 해제, 아니면 새로 구독. 게시판 글 목록 화면에서만 동작.
        """
        if self.current_view != VIEW_POST_LIST or not self.current_board_url:
            speak("이 화면에서는 구독을 토글할 수 없습니다. 게시판 글 목록에서 다시 시도해 주세요.")
            wx.MessageBox(
                "구독은 게시판 글 목록 화면에서만 토글할 수 있습니다.",
                "구독 불가", wx.OK | wx.ICON_INFORMATION, self,
            )
            return

        url = self.current_board_url
        name = re.sub(r"^\s*\d+[\.\)]\s*", "",
                      self.current_menu_name or "게시판").strip()
        mgr = self._ensure_subscription_manager()
        if mgr is None:
            return
        existing = mgr.find(url)
        if existing:
            mgr.remove(url)
            speak(f"{name} 구독을 해제했습니다.")
            wx.MessageBox(f"{name} 구독을 해제했습니다.",
                          "구독 해제", wx.OK | wx.ICON_INFORMATION, self)
        else:
            mgr.add(name, url)
            # 처음 추가된 구독은 seen 이 비어 있으니 백그라운드로 채워둠.
            mgr.initial_fill_async()
            speak(f"{name} 을(를) 구독했습니다. 새 글이 올라오면 알림 센터에 표시됩니다.")
            wx.MessageBox(
                f"{name} 을(를) 구독했습니다.\n"
                "이후 올라오는 새 글이 알림 센터(Ctrl+Shift+N)에 표시됩니다.",
                "구독 추가", wx.OK | wx.ICON_INFORMATION, self,
            )

    def on_open_subscriptions(self, event=None):
        """구독 목록 대화상자 — 간단히 ListBox + 해제 버튼."""
        mgr = self._ensure_subscription_manager()
        if mgr is None:
            return
        if not mgr.items:
            speak("구독 중인 게시판이 없습니다.")
            wx.MessageBox(
                "구독 중인 게시판이 없습니다.\n"
                "게시판 글 목록 화면에서 Ctrl+Shift+S 로 구독을 추가하실 수 있습니다.",
                "구독 없음", wx.OK | wx.ICON_INFORMATION, self,
            )
            return

        dlg = wx.SingleChoiceDialog(
            self,
            "구독 중인 게시판입니다. 해제할 게시판을 선택하고 확인을 누르면 해제됩니다.",
            "구독 목록",
            [f"{s.name}    ({s.url})" for s in mgr.items],
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            sel = dlg.GetSelection()
        finally:
            dlg.Destroy()
        if sel < 0 or sel >= len(mgr.items):
            return
        target = mgr.items[sel]
        ans = wx.MessageBox(f"'{target.name}' 구독을 해제할까요?",
                            "구독 해제 확인",
                            wx.YES_NO | wx.ICON_QUESTION, self)
        if ans == wx.YES:
            mgr.remove(target.url)
            speak("구독을 해제했습니다.")

    def _ensure_subscription_manager(self):
        """구독 매니저를 lazy 생성. 이미 있으면 그대로 반환.

        설정값(`subscription_interval_sec`, `check_subscriptions`) 을 읽어
        타이머 주기를 정한다. 사용 안 함이면 매니저는 만들되 타이머는 안 돈다.
        """
        if self._subscription_manager is not None:
            return self._subscription_manager
        if not self.session:
            return None
        try:
            from subscriptions import SubscriptionManager
            self._subscription_manager = SubscriptionManager(
                self, self.session, self._on_new_subscription_posts,
            )
            self._subscription_manager.initial_fill_async()
            self._restart_subscription_timer()
        except Exception:
            self._subscription_manager = None
        return self._subscription_manager

    def _restart_subscription_timer(self):
        """설정에 따라 구독 폴링 타이머를 재시작/중단."""
        try:
            if getattr(self, "_subscription_timer", None):
                self._subscription_timer.Stop()
        except Exception:
            pass
        try:
            from settings_dialog import load_notify_settings
            settings = load_notify_settings()
            if not bool(settings.get("check_subscriptions", True)):
                return
            interval = max(1, int(settings.get("subscription_interval_sec", 10)))
            self._subscription_timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self._on_subscription_tick,
                      self._subscription_timer)
            self._subscription_timer.Start(interval * 1000)
        except Exception:
            pass

    def _on_subscription_tick(self, event):
        if self._subscription_manager is not None:
            self._subscription_manager.poll_once_async()

    def _on_new_subscription_posts(self, board_name: str, new_posts: list):
        """구독 게시판에 새 글 도착 — 알림 센터 등록 + 사운드/TTS."""
        if not new_posts:
            return
        try:
            from notification import NotificationItem, get_center
            center = get_center()
            to_add = [
                NotificationItem(
                    type="post",
                    item_id=str(getattr(p, "url", "")),
                    sender=board_name,
                    summary=getattr(p, "title", "") or "(제목 없음)",
                    timestamp=getattr(p, "date", ""),
                    extra=p,
                ) for p in new_posts
            ]
            center.add_many(to_add)
        except Exception:
            pass
        # 모달 사용 중이면 소리·TTS 억제 (메일·쪽지 알림과 동일 정책)
        if self._is_modal_dialog_open():
            return
        try:
            from sound import play_event
            play_event("memo_new")
        except Exception:
            pass
        n = len(new_posts)
        if n == 1:
            speak(f"{board_name} 게시판에 새 글이 올라왔습니다. {new_posts[0].title}")
        else:
            speak(f"{board_name} 게시판에 새 글 {n}개가 올라왔습니다.")

    # ── v1.7 DAISY 변환 ──

    def on_convert_daisy(self, event=None):
        """DAISY 도서 ZIP 을 선택해 압축 해제 + TXT 변환."""
        dlg = wx.FileDialog(
            self, "DAISY 도서 ZIP 파일 선택",
            wildcard="ZIP 파일 (*.zip)|*.zip|모든 파일 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            zip_path = dlg.GetPath()
        finally:
            dlg.Destroy()

        from daisy import is_daisy_zip, convert_zip_to_text
        if not is_daisy_zip(zip_path):
            speak("DAISY 도서가 아닌 것 같습니다.")
            ans = wx.MessageBox(
                "선택한 파일이 DAISY 도서로 판정되지 않았습니다.\n"
                "(.opf / .ncx / DTBook XML 이 들어 있어야 합니다)\n\n"
                "그래도 압축 해제하고 변환을 시도할까요?",
                "DAISY 아님", wx.YES_NO | wx.ICON_QUESTION, self,
            )
            if ans != wx.YES:
                return

        speak("DAISY 도서를 변환하는 중입니다.")
        try:
            result = convert_zip_to_text(zip_path)
        except Exception as e:
            speak("DAISY 변환에 실패했습니다.")
            wx.MessageBox(f"변환 중 오류가 발생했습니다.\n{e}",
                          "변환 실패", wx.OK | wx.ICON_ERROR, self)
            return
        if result is None:
            speak("DTBook 본문을 찾지 못해 변환에 실패했습니다.")
            wx.MessageBox(
                "압축은 풀었지만 본문 XML 을 찾지 못해 TXT 변환이 어렵습니다.\n"
                "압축 해제된 폴더를 직접 확인해 주세요.",
                "변환 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return
        folder, txt_path = result
        speak(f"변환 완료. 텍스트 파일은 {os.path.basename(txt_path)} 입니다.")
        ans = wx.MessageBox(
            f"DAISY 도서 변환을 마쳤습니다.\n\n"
            f"폴더: {folder}\n파일: {txt_path}\n\n"
            "폴더를 탐색기로 열어 보시겠습니까?",
            "변환 완료", wx.YES_NO | wx.ICON_INFORMATION, self,
        )
        if ans == wx.YES:
            try:
                os.startfile(folder)
            except Exception:
                pass

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
                # 매 tick 마다 서버 기준 안 읽은 메일 수를 제목 표시줄에 반영.
                self._mail_notifier.on_unread_count = self._set_mail_unread_count
                self._mail_notifier.start_initial_fill()
            if check_memo:
                self._memo_notifier = MemoNotifier(self, self.session, self._on_new_memo_or_mail)
                # 매 tick 마다 서버 기준 안 읽은 쪽지 수를 제목 표시줄에 반영.
                self._memo_notifier.on_unread_count = self._set_memo_unread_count
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

        # v1.7 — 게시판 구독도 함께 시작 (사용자가 구독 항목을 가지고 있을 때만 의미)
        try:
            self._ensure_subscription_manager()
        except Exception:
            pass

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
        """새 메일 도착 — 알림 센터에 등록 + 사운드/TTS + 팝업.

        사용자가 메일함·쪽지함 등 모달 대화상자를 이미 열고 작업 중이면 사운드·
        TTS·팝업은 모두 생략한다. (사용자는 그 화면에서 이미 메일 목록을 보고
        있으므로 별도 알림이 오히려 방해가 된다.) 알림 센터에는 그대로 추가되어
        나중에 확인할 수 있다.
        """
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
        # 모달 대화상자 사용 중이면 모든 알림(소리·TTS·팝업) 생략.
        if self._is_modal_dialog_open():
            return
        try:
            from sound import play_event
            play_event("memo_new")
        except Exception:
            pass
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

    def _is_modal_dialog_open(self) -> bool:
        """현재 자식 모달 대화상자가 열려 있는지."""
        for child in self.GetChildren():
            if isinstance(child, wx.Dialog) and child.IsModal():
                return True
        return False

    def restart_memo_notifier(self):
        """설정 변경 후 호출 — 기존 타이머 중단 후 새 주기로 재시작.

        v1.7 — 게시판 구독 타이머도 같이 재시작한다.
        """
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
        # 구독 타이머도 새 주기로 재시작
        if getattr(self, "_subscription_manager", None):
            self._restart_subscription_timer()

    def _on_new_memo(self, count: int, new_items: list):
        """새 쪽지 도착 콜백 — 알림 센터에 등록 + 사운드·TTS·제목바 업데이트.

        모달 대화상자가 열려 있으면 사운드·TTS·팝업 모두 생략 — 사용자의
        작업을 방해하지 않는다. 알림 센터에는 그대로 등록되어 나중에 확인 가능.
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

        # 모달 대화상자 사용 중이면 이후 알림(소리·TTS·팝업) 모두 생략.
        if self._is_modal_dialog_open():
            return

        # 2. 사운드
        try:
            from sound import play_event
            play_event("memo_new")
        except Exception:
            pass

        # 3. 제목 표시줄의 안 읽은 수는 MemoNotifier.on_unread_count 콜백이
        # 매 tick 마다 서버 기준으로 갱신. 여기서는 누적 증가시키지 않는다.

        # 4. TTS
        sender = new_items[0].counterpart if new_items else "알 수 없음"
        if count == 1:
            speak(f"새 쪽지가 도착했습니다. 보낸 사람 {sender}")
        else:
            speak(f"새 쪽지가 {count}개 도착했습니다.")

        # 5. 확인 대화상자 — Yes 면 알림 센터 오픈.
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

    def _set_memo_unread_count(self, count: int):
        """MemoNotifier 에서 매 tick 마다 서버 기준 '안 읽은 쪽지 수' 전달 시 호출.

        읽거나 삭제해서 서버 상의 안 읽은 수가 줄면 제목 표시줄에도 즉시 반영.
        """
        if getattr(self, "_unread_memo_count", 0) != count:
            self._unread_memo_count = count
            self._update_title_unread()

    def _set_mail_unread_count(self, count: int):
        """MailNotifier 에서 매 tick 마다 서버 기준 '안 읽은 메일 수' 전달 시 호출."""
        if getattr(self, "_unread_mail_count", 0) != count:
            self._unread_mail_count = count
            self._update_title_unread()

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
        # 대화상자 제목에도 숫자 버전을 명시해 스크린리더가 창 제목을 읽을 때
        # 바로 버전을 알 수 있게 한다. 스크린리더가 어색하게 읽는 특수기호
        # (:, →, —, v 접두사)는 피하고 숫자 버전과 자연스러운 한글만 사용.
        super().__init__(
            parent,
            title=f"초록멀티 업데이트 알림 - 새 버전 {info.version}",
            size=(560, 500),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self._info = info

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # 헤더 — 숫자 버전을 크게 표시
        heading = wx.StaticText(
            panel,
            label=f"초록멀티 {info.version} 업데이트가 공개되었습니다.",
        )
        heading.SetFont(make_font(font_size + 3).Bold())

        # 버전 비교 서브라인 — 기호 없이 한글 문장으로
        sub = wx.StaticText(
            panel,
            label=(
                f"현재 버전 {current_version} 에서 "
                f"새 버전 {info.version} 으로 업데이트됩니다."
            ),
        )
        sub.SetFont(make_font(font_size + 1).Bold())

        # 릴리스 이름이 버전 숫자와 다를 때만 간단히 안내. 릴리스 태그는
        # 버전과 중복 정보라 표시하지 않음.
        extra_lines = []
        if info.name and info.name.strip() and info.name.strip() not in (
            info.tag_name, info.version, f"v{info.version}",
            f"초록멀티 {info.version}", f"초록멀티 v{info.version}",
        ):
            extra_lines.append(f"릴리스 이름 {info.name}")
        meta = wx.StaticText(
            panel, label="\n".join(extra_lines) if extra_lines else "",
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
        sizer.Add(meta, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
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
