"""소리샘 사이트에서 메뉴 목록을 자동 감지하는 모듈"""
import requests
from bs4 import BeautifulSoup

from config import SORISEM_BASE_URL


def detect_menus(session: requests.Session) -> list[dict] | None:
    """
    로그인된 세션으로 소리샘 메인 페이지에 접속하여
    메뉴 항목을 자동으로 추출한다.

    Returns:
        [{"name": "메뉴이름", "url": "/bbs/...", "type": "board"}, ...]
        실패 시 None
    """
    try:
        resp = session.get(SORISEM_BASE_URL, timeout=15)
        soup = BeautifulSoup(resp.text, "lxml")
    except Exception:
        return None

    menus = []
    seen_urls = set()

    # 초록등대 동호회를 0번 메뉴로 고정
    green_menu = {
        "name": "초록등대 동호회",
        "url": "/plugin/ar.club/admin.member.php?cl=green",
        "type": "club",
    }
    menus.append(green_menu)
    seen_urls.add(green_menu["url"])

    # 방법 1: 네비게이션 메뉴 영역에서 링크 추출
    nav_selectors = [
        "#menu_pan a",
        "#header_menu a",
        "#menu_top a",
        "nav a",
        ".gnb a",
        ".lnb a",
        ".snb a",
        ".side_menu a",
        "#aside a",
        "#c_menu a",
        ".menu a",
        ".nav a",
        "#gnb a",
    ]

    for selector in nav_selectors:
        links = soup.select(selector)
        for link in links:
            href = link.get("href", "").strip()
            text = link.get_text(strip=True)

            if not text or not href:
                continue
            # 의미 없는 링크 제외
            if href in ("#", "javascript:void(0)", "javascript:;"):
                continue
            if "login" in href or "logout" in href or "register" in href:
                continue
            if len(text) < 2 or len(text) > 30:
                continue

            # 상대경로 정규화
            if href.startswith("http"):
                # 외부 링크 제외
                if SORISEM_BASE_URL not in href:
                    continue
                href = href.replace(SORISEM_BASE_URL, "")

            if href not in seen_urls:
                seen_urls.add(href)
                menu_type = "club" if "ar.club" in href else "board"
                menus.append({
                    "name": text,
                    "url": href,
                    "type": menu_type,
                })

    # 방법 2: bo_table 링크를 모두 찾기 (네비게이션에서 못 찾은 경우)
    if len(menus) <= 1:
        for a_tag in soup.find_all("a", href=True):
            href = a_tag.get("href", "")
            text = a_tag.get_text(strip=True)

            if "bo_table" in href and text and len(text) >= 2:
                if href.startswith("http"):
                    if SORISEM_BASE_URL not in href:
                        continue
                    href = href.replace(SORISEM_BASE_URL, "")

                if href not in seen_urls:
                    seen_urls.add(href)
                    menus.append({
                        "name": text,
                        "url": href,
                        "type": "board",
                    })

    # 방법 3: 동호회 링크도 찾기
    for a_tag in soup.find_all("a", href=True):
        href = a_tag.get("href", "")
        text = a_tag.get_text(strip=True)

        if "ar.club" in href and text and len(text) >= 2:
            if href.startswith("http"):
                href = href.replace(SORISEM_BASE_URL, "")

            if href not in seen_urls:
                seen_urls.add(href)
                menus.append({
                    "name": text,
                    "url": href,
                    "type": "club",
                })

    return menus if len(menus) > 1 else None


def print_detected_menus(menus: list[dict]) -> str:
    """감지된 메뉴를 보기 좋게 포맷한다."""
    lines = []
    for i, m in enumerate(menus):
        lines.append(f"{i}. {m['name']} ({m['url']})")
    return "\n".join(lines)
