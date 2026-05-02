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


# 가상 하위 메뉴 매핑.
# sorisem 응답이 비어 있거나 사이트 구조상 hub 페이지가 없는 메인 메뉴
# 항목에 대해, 코드에서 직접 sub-item 목록을 정의한다. _load_and_show 가
# 해당 URL 을 만나면 fetch 대신 이 목록을 그대로 sub-menu 로 표시한다.
# (이름, url, type) 형식.
VIRTUAL_SUBMENUS: dict[str, tuple[tuple[str, str, str], ...]] = {
    # 7. 전자도서관 — hub 가 빈 ul 을 응답해 직접 정의.
    "/?mo=lib2013&cl=lib2013": (
        ("1. 소설",        "/?mo=novs&cl=lib2013",                     "category"),
        ("2. 시/에세이",   "/?mo=poes&cl=lib2013",                     "category"),
        ("3. 경제 / 경영", "/bbs/board.php?bo_table=eco&cl=lib2013",   "board"),
        ("4. 정치 / 사회", "/bbs/board.php?bo_table=soci&cl=lib2013",  "board"),
        ("5. 인문",        "/bbs/board.php?bo_table=hum&cl=lib2013",   "board"),
        ("6. 역사",        "/bbs/board.php?bo_table=his&cl=lib2013",   "board"),
        ("7. 과학 / 기술 / IT", "/bbs/board.php?bo_table=sci&cl=lib2013", "board"),
        ("8. 건강 / 심리", "/bbs/board.php?bo_table=hea&cl=lib2013",   "board"),
        ("9. 국어 / 외국어", "/bbs/board.php?bo_table=lan&cl=lib2013", "board"),
        ("10. 유아 / 어린이 / 청소년", "/bbs/board.php?bo_table=chi&cl=lib2013", "board"),
        ("11. 종교",       "/bbs/board.php?bo_table=rel&cl=lib2013",    "board"),
        ("12. 기타(예술, 대중문화, 가정, 취미)", "/bbs/board.php?bo_table=etcs&cl=lib2013", "board"),
        ("55. 전자도서 신청란", "/bbs/board.php?bo_table=booksub2&cl=lib2013", "board"),
        ("77. 공지사항",   "/bbs/board.php?bo_table=alllib4&cl=lib2013", "board"),
        ("88. 전자도서 문의", "/bbs/board.php?bo_table=alllib5&cl=lib2013", "board"),
        ("99. 전체자료실", "/bbs/board.php?bo_table=alllib99&cl=lib2013", "board"),
    ),
    # 7-1. 소설 — sub-category. 같은 hub URL 패턴이라 자동 파싱이 어려워
    # 사용자가 알려준 항목을 직접 정의.
    "/?mo=novs&cl=lib2013": (
        ("1. 일반소설",     "/bbs/board.php?bo_table=aetcs&cl=lib2013",    "board"),
        ("2. 로멘스",       "/bbs/board.php?bo_table=romances&cl=lib2013", "board"),
        ("3. 무협 / 판타지", "/bbs/board.php?bo_table=chils&cl=lib2013",    "board"),
        ("4. 추리 / 스릴러", "/bbs/board.php?bo_table=detes&cl=lib2013",    "board"),
        ("5. 역사소설",     "/bbs/board.php?bo_table=wars&cl=lib2013",     "board"),
    ),
    # 7-2. 시/에세이 — sub-category.
    "/?mo=poes&cl=lib2013": (
        ("1. 시",     "/bbs/board.php?bo_table=poe&cl=lib2013",  "board"),
        ("2. 에세이", "/bbs/board.php?bo_table=essa&cl=lib2013", "board"),
    ),
}


# 표준 소리샘 메인 메뉴 — 자동 감지나 사용자 편집에서 빠지면 자동 보충.
# (이름, url, type) 순서. 맨 앞 숫자 접두사가 정렬 기준.
# v1.7 — 6.자료실, 7.전자도서관 의 URL 을 cl= 가 함께 있는 형태로 수정.
# /?mo=pds 등 단독 mo URL 로는 sorisem 이 메인 사이드바만 응답해 빈 카테고리로 보였다.
WELL_KNOWN_MAIN_MENUS: tuple[tuple[str, str, str], ...] = (
    ("초록등대 동호회", "/plugin/ar.club/?cl=green", "club"),
    ("1. 소리샘 공지사항", "/bbs/board.php?bo_table=sorisemnotice", "board"),
    ("3. 개발자 포럼", "/?mo=prg", "category"),
    ("4. 동호회", "/?mo=potion", "category"),
    ("5. 잡지", "/?mo=magazin", "category"),
    ("6. 자료실", "/?mo=pds&cl=pds", "category"),
    # 7번 전자도서관 — sorisem 의 hub URL 이 빈 ul 을 응답해 자동 파싱 불가.
    # VIRTUAL_SUBMENUS 에 정의된 가상 하위 메뉴를 사용한다.
    ("7. 전자도서관", "/?mo=lib2013&cl=lib2013", "category"),
    ("8. 노원시각장애인학습지원센터", "/?mo=edu2013&cl=edu2013", "category"),
    ("9. 점자도서관", "/?mo=braille", "category"),
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
                # 사용자 txt 에 빠진 표준 메뉴(예: 6.자료실, 7.전자도서관,
                # 8.노원시각장애인학습지원센터) 가 있으면 자동 보충 + 정돈.
                self._ensure_forced_club_menus()
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

        # 3) 표준 메인 메뉴 항목 자동 보충 — 사용자 편집·자동 감지에서 빠진
        #    잘 알려진 항목(예: 6.자료실, 7.전자도서관, 8.노원시각장애인학습지원
        #    센터)을 다시 채워 넣어 사용자가 사이트 기본 메뉴를 그대로 쓸 수 있게 한다.
        existing_urls = {m.url for m in self.menus}
        if self._supplement_well_known(existing_urls):
            changed = True

        # 4) 같은 이름으로 중복된 엔트리 제거. 예전 버전이 "6. 자료실" 을
        #    "6. 소리샘 자료실" 로 이름 보정한 직후 supplement 가 새 URL
        #    (/?mo=pds&cl=pds) 의 "6. 자료실" 을 추가했고, 다음 로딩에서 그것까지
        #    "6. 소리샘 자료실" 로 다시 보정되어 동명 엔트리 두 개가 남는 사례가
        #    발생했다. WELL_KNOWN URL 을 우선 보존, 없으면 마지막 항목을 남긴다.
        from collections import OrderedDict
        well_known_urls = {url for _n, url, _t in WELL_KNOWN_MAIN_MENUS}
        by_name: OrderedDict[str, MenuItem] = OrderedDict()
        for m in self.menus:
            prev = by_name.get(m.name)
            if prev is None:
                by_name[m.name] = m
                continue
            # 동명 중복 — WELL_KNOWN 에 등록된 URL 을 우선 보존.
            if m.url in well_known_urls and prev.url not in well_known_urls:
                by_name[m.name] = m
        if len(by_name) != len(self.menus):
            self.menus = list(by_name.values())
            changed = True

        return changed

    @staticmethod
    def _url_match_key(url: str) -> str:
        """매뉴얼/자동감지 URL 을 동일성 비교용 정규화 형태로 변환.

        sorisem 은 같은 카테고리를 두 가지 URL 로 노출한다:
          · 자동감지 결과: `/?mo=pds`
          · WELL_KNOWN: `/?mo=pds&cl=pds`
        둘 다 동일한 자료실인데 supplement 가 별개로 인식해 중복 엔트리를
        만드는 문제가 있었다. mo= 값과 board.php 의 bo_table 값을 키로 추출해
        같은 카테고리·게시판이면 같은 키를 갖도록 한다.
        """
        if not url:
            return ""
        u = url.lower().strip()
        m_mo = re.search(r"[?&]mo=([a-z0-9_]+)", u)
        if m_mo:
            return f"mo={m_mo.group(1)}"
        m_bo = re.search(r"[?&]bo_table=([a-z0-9_]+)", u)
        if m_bo:
            return f"bo={m_bo.group(1)}"
        return u

    def _supplement_well_known(self, existing_urls: set[str]) -> bool:
        """`WELL_KNOWN_MAIN_MENUS` 중 self.menus 에 빠진 항목을
        번호 순서대로 적절한 위치에 끼워 넣는다.

        매칭 우선순위:
          · 같은 이름(번호 접두사 포함) 의 항목이 이미 있으면 URL 만 갱신.
          · URL 정규화 키(예: mo=pds) 가 같은 항목이 있으면 URL 갱신만 하고
            새로 추가하지 않음 — 자동감지 `/?mo=pds` 와 WELL_KNOWN
            `/?mo=pds&cl=pds` 같은 중복 인식 문제 방지.
          · 같은 URL 이 이미 있으면 건너뜀.
          · 위 어느 매칭도 없을 때만 새로 삽입.

        반환: 실제로 변경(추가·URL 갱신) 이 있었는지 여부.
        """
        changed = False
        existing_by_name: dict[str, MenuItem] = {}
        existing_by_key: dict[str, MenuItem] = {}
        for m in self.menus:
            existing_by_name[m.name] = m
            key = self._url_match_key(m.url)
            if key and key not in existing_by_key:
                existing_by_key[key] = m
        for name, url, mtype in WELL_KNOWN_MAIN_MENUS:
            existing = existing_by_name.get(name)
            if existing is not None:
                if existing.url != url:
                    existing.url = url
                    existing.type = mtype
                    changed = True
                continue
            # URL 정규화 키 매칭 — 자동감지가 다른 이름으로 같은 카테고리를
            # 등록한 경우(예: "6. 소리샘 자료실" /?mo=pds) WELL_KNOWN URL
            # (/?mo=pds&cl=pds) 으로 보정만 하고 새 엔트리 추가는 막는다.
            key = self._url_match_key(url)
            existing_by_key_match = existing_by_key.get(key) if key else None
            if existing_by_key_match is not None:
                if existing_by_key_match.url != url:
                    existing_by_key_match.url = url
                    existing_by_key_match.type = mtype
                    changed = True
                continue
            if url in existing_urls:
                continue
            new_item = MenuItem(name=name, url=url, menu_type=mtype)
            insert_at = self._find_insert_index(name)
            self.menus.insert(insert_at, new_item)
            existing_urls.add(url)
            existing_by_name[name] = new_item
            if key:
                existing_by_key[key] = new_item
            changed = True
        return changed

    def _find_insert_index(self, name: str) -> int:
        """이름의 숫자 접두사를 보고 self.menus 의 정렬을 유지하는 위치를 찾는다.

        숫자 접두사가 없으면 끝에 추가.
        """
        m = re.match(r'^\s*(\d+)[\.\)]\s*', name or "")
        if not m:
            return len(self.menus)
        my_num = int(m.group(1))
        for idx, mi in enumerate(self.menus):
            mm = re.match(r'^\s*(\d+)[\.\)]\s*', mi.name or "")
            if not mm:
                continue
            other_num = int(mm.group(1))
            if my_num < other_num:
                return idx
        return len(self.menus)

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
