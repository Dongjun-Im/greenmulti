"""메뉴 목록 관리 모듈"""
import json
import os
import re
import unicodedata
from urllib.parse import urlparse

from config import MENU_LIST_FILE, MENU_LIST_TXT_FILE, SORISEM_BASE_URL


MENU_TXT_HEADER = """\
# 초록멀티 메뉴 목록 — 사용자 편집 파일
#
# 이 파일을 수정하면 프로그램이 자동으로 반영합니다.
# (이 파일이 존재하면 자동 메뉴 감지 결과보다 우선합니다.)
# 파일을 삭제하면 로그인 시 소리샘 메인 페이지에서 메뉴를 자동으로 다시 긁어옵니다.
#
# 한 줄에 하나씩, 아래 형식으로 적습니다.
#
#   이름 | URL | 타입
#
# - 이름: 메뉴에 표시할 이름
# - URL: /plugin/ar.club/?cl=green 처럼 소리샘 기준 상대 경로
#         또는 http(s):// 로 시작하는 절대 URL
# - 타입(선택): board, club, category, external 중 하나. 생략 가능.
#
# '#' 으로 시작하는 줄, 빈 줄은 무시됩니다.
"""


def _infer_type(url: str) -> str:
    """URL에서 메뉴 타입 추론."""
    if not url:
        return "board"
    u = url.lower()
    if u.startswith("http") and SORISEM_BASE_URL.lower() not in u:
        return "external"
    if "ar.club" in u:
        return "club"
    if "bo_table" in u:
        return "board"
    return "category"


# 초록등대 동호회 하위 메뉴에서 "자료실" / "엔터테인먼트 자료실" 이름을 가진
# 게시판에 대해 바로가기 코드를 강제하기 위한 매핑표.
# (메인 메뉴에 추가하는 용도로는 더 이상 사용하지 않는다. 메인 메뉴에는 소리샘
# 자체의 자료실 /?mo=pds 만 "소리샘 자료실" 이름으로 남긴다.)
FORCED_CLUB_MENUS: tuple[tuple[str, str], ...] = (
    ("자료실", "/plugin/ar.club/?cl=green4"),
    ("엔터테인먼트 자료실", "/plugin/ar.club/?cl=green6"),
)


def _core_name(name: str) -> str:
    """메뉴 이름에서 앞쪽 번호 접두사(예: "4. ")와 공백을 제거한 핵심 이름.

    저장된 메뉴 이름이 "4. 자료실" 처럼 번호가 붙어 저장되어 있는 경우가 있어
    강제 보정 로직이 이름을 매칭할 때 이 함수를 통해 접두사를 벗겨 비교한다.
    """
    return re.sub(r'^\s*\d+[\.\)]\s*', '', name or '').strip()


def _forced_shortcut_code(name: str) -> str:
    """메뉴 이름이 정확히 "자료실" / "엔터테인먼트 자료실" 이면 강제 바로가기 코드.

    번호 접두사("4. ")와 앞뒤 공백만 허용. "일반자료실", "포터블자료실",
    "기타자료실" 같이 다른 글자와 결합된 이름은 매치하지 않아 본래 게시판
    코드(green42 등)를 덮어쓰지 않는다. 초록등대 자료실(NAS)은 배제.

    macOS 파일 시스템은 한글을 자모 분리(NFD)로 저장하므로 NFC 로 통일 후 비교.

    반환 "" 이면 강제 매핑 없음(일반 URL 기반 코드 사용).
    """
    s = unicodedata.normalize('NFC', name or '').strip()
    # 번호 접두사 제거 ("4. 자료실" -> "자료실")
    s = re.sub(r'^\s*\d+[\.\)]\s*', '', s).strip()

    # 정확 일치만 인정 — 단어 일부 포함은 거부
    if s == unicodedata.normalize('NFC', '엔터테인먼트 자료실'):
        return 'green6'
    if s == unicodedata.normalize('NFC', '자료실'):
        return 'green4'
    return ''


def extract_shortcut_code(url: str) -> str:
    """URL에서 바로가기 코드(고유 식별자)를 추출한다.

    - 외부 URL: 도메인의 주 부분 (예: youtube.com → youtube)
    - cl=xxx 파라미터: xxx
    - bo_table=xxx 파라미터: xxx
    - mo=xxx 파라미터: xxx
    - /xxx/ 경로: xxx
    """
    if not url:
        return ""

    # 외부 URL
    if url.startswith("http"):
        host = urlparse(url).hostname or ""
        host = host.replace("www.", "")
        parts = host.split(".")
        if parts:
            return parts[0]
        return ""

    # bo_table=xxx 우선 (게시판이 가장 고유함)
    m = re.search(r'bo_table=([a-zA-Z0-9_]+)', url)
    if m:
        return m.group(1)

    # cl=xxx (클럽)
    m = re.search(r'cl=([a-zA-Z0-9_]+)', url)
    if m:
        return m.group(1)

    # mo=xxx
    m = re.search(r'mo=([a-zA-Z0-9_]+)', url)
    if m:
        return m.group(1)

    # /mypage/ 같은 경로 기반
    m = re.search(r'/([a-zA-Z0-9_]+)/?(?:\?|$)', url)
    if m:
        return m.group(1)

    return ""


class MenuItem:
    """메뉴 항목"""

    def __init__(self, name: str, url: str, menu_type: str = "board"):
        self.name = name
        self.url = url
        self.type = menu_type  # board, club, category, external

    @property
    def full_url(self) -> str:
        if self.url.startswith("http"):
            return self.url
        return f"{SORISEM_BASE_URL}{self.url}"

    @property
    def is_external(self) -> bool:
        return self.type == "external"

    @property
    def shortcut_code(self) -> str:
        """바로가기 코드. 자료실·엔터테인먼트 자료실은 이름 기반으로 강제."""
        forced = _forced_shortcut_code(self.name)
        if forced:
            return forced
        return extract_shortcut_code(self.url)

    def __str__(self) -> str:
        return self.name


class MenuManager:
    """메뉴 목록 로드/저장 관리"""

    def __init__(self):
        self.menus: list[MenuItem] = []

    def load(self) -> list[MenuItem]:
        # 사용자가 직접 편집한 텍스트 파일이 있으면 최우선으로 사용.
        # (자동 감지 결과보다 우선하며, 자동 감지는 이 파일을 덮어쓰지 않는다.)
        if os.path.exists(MENU_LIST_TXT_FILE):
            txt_menus = self._load_from_txt()
            if txt_menus:
                self.menus = txt_menus
                # JSON 캐시도 동기화해 다른 코드 경로와 일관성 유지
                try:
                    self.save()
                except Exception:
                    pass
                return self.menus

        if not os.path.exists(MENU_LIST_FILE):
            self.menus = self._default_menus()
            self.save()
            return self.menus

        try:
            with open(MENU_LIST_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.menus = []
            for item in data.get("menus", []):
                self.menus.append(MenuItem(
                    name=item["name"],
                    url=item["url"],
                    menu_type=item.get("type", "board"),
                ))
        except (json.JSONDecodeError, KeyError):
            self.menus = self._default_menus()

        # 메인 메뉴에서 "nas" 타입 엔트리 제거 — 이제 메뉴바 '도구' 메뉴로 이동함
        before = len(self.menus)
        self.menus = [m for m in self.menus if m.type != "nas"]

        # 자료실 / 엔터테인먼트 자료실을 강제로 보장 (빠져 있으면 추가, URL 보정)
        changed = self._ensure_forced_club_menus()

        if len(self.menus) != before or changed:
            try:
                self.save()
            except Exception:
                pass
        return self.menus

    def _ensure_forced_club_menus(self) -> bool:
        """메인 메뉴의 자료실 엔트리를 "소리샘 자료실"로 정돈.

        - 이전 버전에서 메인 메뉴에 잘못 추가된 초록등대 클럽 자료실
          (cl=green4) / 엔터테인먼트 자료실 (cl=green6) 엔트리는 제거.
          이 두 메뉴는 초록등대 동호회 하위 메뉴에만 남겨 둔다.
        - 소리샘 자체의 자료실(URL에 mo=pds 포함)은 이름에 번호 접두사를
          유지한 채 "소리샘 자료실"로 표시되도록 보정.

        반환: 실제 변경이 있었는지 여부.
        """
        changed = False

        # 1) 이전 버전에서 추가된 초록등대 클럽 자료실/엔터 자료실 엔트리 제거
        STALE_URLS = {
            "/plugin/ar.club/?cl=green4",
            "/plugin/ar.club/?cl=green6",
        }
        before = len(self.menus)
        self.menus = [m for m in self.menus if m.url not in STALE_URLS]
        if len(self.menus) != before:
            changed = True

        # 2) 소리샘 자료실 이름 보정: "자료실"이라는 이름(번호 접두사 허용)이
        #    /?mo=pds URL에 있으면 "소리샘 자료실"로 변경
        for m in self.menus:
            if "mo=pds" not in m.url:
                continue
            m_num = re.match(r'^(\s*\d+[\.\)]\s*)', m.name or "")
            prefix = m_num.group(1) if m_num else ""
            core = (m.name or "")[len(prefix):].strip()
            if core == "자료실":
                new_name = f"{prefix}소리샘 자료실"
                if m.name != new_name:
                    m.name = new_name
                    changed = True

        return changed

    def save(self) -> None:
        data = {
            "version": 3,
            "description": "초록멀티 메뉴 목록",
            "menus": [
                {"name": m.name, "url": m.url, "type": m.type}
                for m in self.menus
            ],
        }
        os.makedirs(os.path.dirname(MENU_LIST_FILE), exist_ok=True)
        with open(MENU_LIST_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def has_user_override(self) -> bool:
        """사용자가 직접 편집한 텍스트 메뉴 파일이 존재하는지 여부."""
        return os.path.exists(MENU_LIST_TXT_FILE)

    def _load_from_txt(self) -> list[MenuItem]:
        """사용자 편집용 텍스트 메뉴 파일 파서.

        포맷: 한 줄에 "이름 | URL | 타입(선택)". '#' 주석·빈 줄 허용.
        파싱 실패하거나 유효 항목이 0개면 빈 리스트 반환(자동 감지로 폴백).
        """
        menus: list[MenuItem] = []
        try:
            with open(MENU_LIST_TXT_FILE, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except OSError:
            return []

        for raw in lines:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 2:
                continue
            name, url = parts[0], parts[1]
            if not name or not url:
                continue
            menu_type = parts[2] if len(parts) >= 3 and parts[2] else _infer_type(url)
            menus.append(MenuItem(name=name, url=url, menu_type=menu_type))
        return menus

    def export_to_txt(self) -> None:
        """현재 메뉴 목록을 사용자 편집용 텍스트 파일로 저장.

        파일이 이미 존재하면 덮어쓰지 않고(사용자 편집 보존) 그대로 둔다.
        """
        if os.path.exists(MENU_LIST_TXT_FILE):
            return
        os.makedirs(os.path.dirname(MENU_LIST_TXT_FILE), exist_ok=True)
        lines = [MENU_TXT_HEADER]
        for m in self.menus:
            lines.append(f"{m.name} | {m.url} | {m.type}")
        with open(MENU_LIST_TXT_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")

    def get_display_names(self) -> list[str]:
        """메인 메뉴 표시 목록 (바로가기 코드 포함)"""
        items = []
        for i, m in enumerate(self.menus):
            # 번호 접두사
            if re.match(r'^\d+[\.\)]\s', m.name):
                base = m.name
            else:
                base = f"{i}. {m.name}"
            # 바로가기 코드 — 자료실·엔터테인먼트 자료실은 URL 이 어떻든
            # 항상 green4 / green6 로 표시되도록 강제한다.
            code = self._resolve_display_code(m)
            if code:
                items.append(f"{base} (바로가기 코드: {code})")
            else:
                items.append(base)
        return items

    def get_shortcut_codes(self) -> list[str]:
        """각 메뉴의 바로가기 코드 목록 (표시 순서와 동일)"""
        return [self._resolve_display_code(m) for m in self.menus]

    @staticmethod
    def _resolve_display_code(m: "MenuItem") -> str:
        """표시용 바로가기 코드. 강제 메뉴면 항상 고정 코드를 반환."""
        return m.shortcut_code

    def get_menu_by_index(self, index: int) -> MenuItem | None:
        if 0 <= index < len(self.menus):
            return self.menus[index]
        return None

    def _default_menus(self) -> list[MenuItem]:
        return [
            MenuItem("초록등대 동호회", "/plugin/ar.club/?cl=green", "club"),
            MenuItem("소리샘 공지사항", "/bbs/board.php?bo_table=sorisemnotice", "board"),
        ]
