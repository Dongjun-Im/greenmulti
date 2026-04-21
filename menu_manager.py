"""메뉴 목록 관리 모듈"""
import json
import os
import re
from urllib.parse import urlparse

from config import MENU_LIST_FILE, SORISEM_BASE_URL


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
        """URL 기반 바로가기 코드"""
        return extract_shortcut_code(self.url)

    def __str__(self) -> str:
        return self.name


class MenuManager:
    """메뉴 목록 로드/저장 관리"""

    def __init__(self):
        self.menus: list[MenuItem] = []

    def load(self) -> list[MenuItem]:
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
            return self.menus

        except (json.JSONDecodeError, KeyError):
            self.menus = self._default_menus()
            return self.menus

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

    def get_display_names(self) -> list[str]:
        """메인 메뉴 표시 목록 (바로가기 코드 포함)"""
        items = []
        for i, m in enumerate(self.menus):
            # 번호 접두사
            if re.match(r'^\d+[\.\)]\s', m.name):
                base = m.name
            else:
                base = f"{i}. {m.name}"
            # 바로가기 코드 추가
            code = m.shortcut_code
            if code:
                items.append(f"{base} (바로가기 코드: {code})")
            else:
                items.append(base)
        return items

    def get_shortcut_codes(self) -> list[str]:
        """각 메뉴의 바로가기 코드 목록 (표시 순서와 동일)"""
        return [m.shortcut_code for m in self.menus]

    def get_menu_by_index(self, index: int) -> MenuItem | None:
        if 0 <= index < len(self.menus):
            return self.menus[index]
        return None

    def _default_menus(self) -> list[MenuItem]:
        return [
            MenuItem("초록등대 동호회", "/plugin/ar.club/?cl=green", "club"),
            MenuItem("소리샘 공지사항", "/bbs/board.php?bo_table=sorisemnotice", "board"),
        ]
