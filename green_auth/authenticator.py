"""소리샘 로그인 및 초록등대 동호회 인증 모듈"""
import re

import requests
from bs4 import BeautifulSoup

from green_auth.config import LOGIN_URL, SORISEM_BASE_URL, GREEN_CLUB_MEMBERS_URL


# 초록등대 동호회 인증 허용 등급
# - 등급 6: 일반 회원
# - 등급 7: 우수 회원
# - 등급 8: 최우수 회원
# - 등급 9: 명예 회원
# - 동호회관리자
ALLOWED_RANK_PATTERNS = [
    re.compile(r"동호회\s*관리자"),
    re.compile(r"클럽\s*관리자"),
    re.compile(r"최우수\s*회원"),
    re.compile(r"명예\s*회원"),
    re.compile(r"우수\s*회원"),
    re.compile(r"일반\s*회원"),
]


def _find_allowed_rank(text: str) -> str | None:
    """텍스트에서 허용된 등급 텍스트를 찾는다."""
    for pattern in ALLOWED_RANK_PATTERNS:
        match = pattern.search(text)
        if match:
            return match.group(0)
    return None


class AuthResult:
    """인증 결과"""
    SUCCESS = "success"
    LOGIN_FAILED = "login_failed"
    NOT_MEMBER = "not_member"
    NETWORK_ERROR = "network_error"

    def __init__(self, status: str, message: str = ""):
        self.status = status
        self.message = message

    @property
    def is_success(self) -> bool:
        return self.status == self.SUCCESS


class Authenticator:
    """소리샘 로그인 + 초록등대 동호회 등급 인증"""

    # 회원 목록 후보 URL: 관리자 페이지가 일반 회원에게 막힐 수 있어 다중 시도
    MEMBER_PAGE_URLS = [
        f"{SORISEM_BASE_URL}/plugin/ar.club/admin.member.php?cl=green",
        f"{SORISEM_BASE_URL}/plugin/ar.club/member.php?cl=green",
        f"{SORISEM_BASE_URL}/plugin/ar.club/member_list.php?cl=green",
        f"{SORISEM_BASE_URL}/plugin/ar.club/list.php?cl=green",
        f"{SORISEM_BASE_URL}/plugin/ar.club/?cl=green",
    ]

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "GreenAuth/1.0",
        })
        # 인증 성공한 사용자의 소리샘 로그인 아이디.
        # MainFrame 이 본인 게시물 여부 검증 등에 사용한다.
        self.user_id: str | None = None
        # 로그인 사용자의 닉네임 (게시물 작성자 일치 여부 검증용).
        self.nickname: str | None = None

    def authenticate(self, user_id: str, password: str) -> AuthResult:
        """소리샘 로그인 후 초록등대 동호회 회원 등급을 확인한다."""
        login_result = self._login(user_id, password)
        if not login_result.is_success:
            return login_result

        result = self._check_green_membership(user_id)
        if result.is_success:
            self.user_id = user_id
            self.nickname = self._fetch_my_nickname(user_id)
        return result

    # ─────────────────────────── 닉네임 조회 ───────────────────────────

    def _fetch_my_nickname(self, user_id: str) -> str | None:
        """로그인 사용자의 소리샘 닉네임을 조회. 실패해도 앱 실행엔 지장 없도록
        None 반환. 이후 본인 게시물 여부 검증에서 닉네임이 없으면 다른 전략으로
        폴백된다.

        gnuboard 계열은 `/bbs/profile.php?mb_id=X` 가 사용자 프로필 페이지를
        돌려주므로 여기서 닉네임을 추출하는 것이 가장 안정적. 실패 시 메인
        페이지의 로그인 영역을 2차로 탐색."""
        nick = self._fetch_nickname_from_profile(user_id)
        if nick:
            return nick
        return self._fetch_nickname_from_main()

    def _fetch_nickname_from_profile(self, user_id: str) -> str | None:
        try:
            url = f"{SORISEM_BASE_URL}/bbs/profile.php?mb_id={user_id}"
            resp = self.session.get(url, timeout=15)
        except Exception:
            return None
        if not resp.ok:
            return None

        try:
            soup = BeautifulSoup(resp.text, "lxml")
        except Exception:
            return None

        # 1) 일반적인 프로필 닉네임 선택자
        for sel in [
            ".mb_nick", ".prof_nick", ".nickname",
            "#profile_nick", ".profile_nick",
            ".profile-name", ".prof-name",
            ".sv_name", ".if_name", ".mb_name",
        ]:
            el = soup.select_one(sel)
            if el:
                text = self._clean_nickname(el.get_text(strip=True))
                if text:
                    return text

        # 2) 페이지 title: "XXX님의 정보" 패턴
        if soup.title:
            m = re.search(
                r"([^\s<>]{1,30})\s*님(?:의)?\s*(?:정보|프로필|페이지)?",
                soup.title.get_text(strip=True),
            )
            if m:
                return self._clean_nickname(m.group(1))

        # 3) h1/h2 헤더
        for sel in ("h1", "h2", "h3"):
            el = soup.select_one(sel)
            if el:
                text = el.get_text(strip=True)
                m = re.search(r"([^\s<>]{1,30})\s*님", text)
                if m:
                    return self._clean_nickname(m.group(1))
                # "닉네임 프로필" 같은 패턴
                if 1 < len(text) < 30 and not text.lower().startswith(
                    ("profile", "정보")
                ):
                    return self._clean_nickname(text)
        return None

    def _fetch_nickname_from_main(self) -> str | None:
        try:
            resp = self.session.get(SORISEM_BASE_URL, timeout=15)
        except Exception:
            return None
        if not resp.ok:
            return None

        html = resp.text
        # 로그인 상태에서 헤더/상단에 "XXX님 환영합니다" 또는 "XXX님" + 근처 로그아웃
        # 링크가 표시되는 일반 gnuboard 패턴을 탐색.
        for pattern in [
            r"([^\s<>\n]{1,30})\s*님[^<>]{0,50}?(?:환영|안녕|반갑|hello)",
            r"([^\s<>\n]{1,30})\s*님[\s\S]{0,250}?로그아웃",
            r"로그아웃[\s\S]{0,250}?([^\s<>\n]{1,30})\s*님",
        ]:
            m = re.search(pattern, html)
            if m:
                cand = self._clean_nickname(m.group(1))
                if cand:
                    return cand
        return None

    @staticmethod
    def _clean_nickname(text: str) -> str:
        if not text:
            return ""
        text = re.sub(r"<[^>]+>", "", text).strip()
        for suffix in ("님의 정보", "님 정보", "의 정보", "의정보", "님"):
            if text.endswith(suffix):
                text = text[: -len(suffix)].strip()
        text = re.sub(r"\s+", " ", text)
        if 1 < len(text) < 30:
            return text
        return ""

    def _login(self, user_id: str, password: str) -> AuthResult:
        """소리샘 사이트 로그인"""
        try:
            self.session.get(SORISEM_BASE_URL, timeout=15)

            login_data = {
                "mb_id": user_id,
                "mb_password": password,
            }
            self.session.post(
                LOGIN_URL,
                data=login_data,
                timeout=15,
                allow_redirects=True,
            )

            main_resp = self.session.get(SORISEM_BASE_URL, timeout=15)
            page_text = main_resp.text

            if "login_check.php" in page_text and "mb_password" in page_text:
                return AuthResult(
                    AuthResult.LOGIN_FAILED,
                    "아이디 또는 비밀번호가 올바르지 않습니다."
                )

            return AuthResult(AuthResult.SUCCESS)

        except requests.exceptions.ConnectionError:
            return AuthResult(
                AuthResult.NETWORK_ERROR,
                "인터넷 연결을 확인해 주세요. 네트워크에 연결할 수 없습니다."
            )
        except requests.exceptions.Timeout:
            return AuthResult(
                AuthResult.NETWORK_ERROR,
                "서버 응답 시간이 초과되었습니다. 잠시 후 다시 시도해 주세요."
            )
        except requests.exceptions.RequestException as e:
            return AuthResult(
                AuthResult.NETWORK_ERROR,
                f"네트워크 오류가 발생했습니다: {e}"
            )

    def _check_green_membership(self, user_id: str) -> AuthResult:
        """초록등대 동호회 회원 목록에서 사용자의 등급을 확인한다.

        허용 등급:
        - 등급 6: 일반 회원
        - 등급 7: 우수 회원
        - 등급 8: 최우수 회원
        - 등급 9: 명예 회원
        - 동호회관리자
        위 등급 중 하나여야 인증 성공. 그 외 등급/비회원은 실패.
        """
        try:
            user_found = False
            matched_rank: str | None = None
            page_full_text = ""

            for base_url in self.MEMBER_PAGE_URLS:
                page = 1
                max_pages = 50
                while page <= max_pages:
                    sep = "&" if "?" in base_url else "?"
                    page_url = base_url if page == 1 else f"{base_url}{sep}page={page}"

                    try:
                        resp = self.session.get(page_url, timeout=15)
                    except requests.exceptions.RequestException:
                        break

                    if not resp.ok:
                        break

                    soup = BeautifulSoup(resp.text, "lxml")
                    user_row = self._find_user_row(soup, user_id)

                    if user_row is not None:
                        user_found = True
                        matched_rank = _find_allowed_rank(user_row)
                        # 행에 등급 텍스트가 없으면 같은 페이지 전체에서도 검색
                        if not matched_rank:
                            page_full_text = soup.get_text(" ", strip=True)
                            matched_rank = _find_allowed_rank(page_full_text)
                        break

                    if not self._has_next_page(soup, page):
                        break
                    page += 1

                if user_found:
                    break

            if not user_found:
                return AuthResult(
                    AuthResult.NOT_MEMBER,
                    "초록등대 동호회 회원이 아닙니다. 프로그램을 종료합니다."
                )

            # 사용자는 회원 목록에 있음. 등급 검증.
            if matched_rank:
                return AuthResult(
                    AuthResult.SUCCESS,
                    f"초록등대 동호회 ({matched_rank}) 인증에 성공했습니다."
                )

            # 회원이지만 등급 텍스트를 추출하지 못한 경우 - 회원 자체는 확인됨
            return AuthResult(
                AuthResult.SUCCESS,
                "초록등대 동호회 회원 인증에 성공했습니다."
            )

        except requests.exceptions.RequestException:
            return AuthResult(
                AuthResult.NETWORK_ERROR,
                "회원 목록을 확인할 수 없습니다. 인터넷 연결을 확인해 주세요."
            )

    def _find_user_row(self, soup: BeautifulSoup, user_id: str) -> str | None:
        """페이지에서 사용자의 행/항목 텍스트를 찾는다. 없으면 None."""
        user_id_lower = user_id.lower()
        word_pattern = re.compile(rf"\b{re.escape(user_id_lower)}\b")

        # 1. 테이블 행
        for tr in soup.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            if not cells:
                continue
            row_text = " ".join(c.get_text(" ", strip=True) for c in cells)
            if word_pattern.search(row_text.lower()):
                return row_text

        # 2. 리스트 항목
        for li in soup.find_all("li"):
            item_text = li.get_text(" ", strip=True)
            if word_pattern.search(item_text.lower()):
                return item_text

        # 3. 회원 카드 (div)
        for div in soup.find_all("div", class_=re.compile(r"member|user|mb_", re.I)):
            div_text = div.get_text(" ", strip=True)
            if word_pattern.search(div_text.lower()):
                return div_text

        # 4. 페이지 전체 텍스트에서 단어 단위로 발견
        full_text = soup.get_text(" ", strip=True)
        if word_pattern.search(full_text.lower()):
            return full_text

        return None

    def _has_next_page(self, soup: BeautifulSoup, current_page: int) -> bool:
        """다음 페이지가 존재하는지 확인"""
        next_page = current_page + 1
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if f"page={next_page}" in href or f"page%3D{next_page}" in href:
                return True
        return False

    def logout(self) -> None:
        """소리샘 로그아웃"""
        try:
            from green_auth.config import LOGOUT_URL
            self.session.get(LOGOUT_URL, timeout=10)
        except Exception:
            pass
        finally:
            self.session.cookies.clear()
