"""소리샘 로그인 및 초록등대 동호회 인증 모듈"""
import requests
from bs4 import BeautifulSoup

from config import LOGIN_URL, SORISEM_BASE_URL, GREEN_CLUB_MEMBERS_URL


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
    """소리샘 로그인 + 초록등대 동호회 인증"""

    def __init__(self):
        self.session = requests.Session()
        # 브라우저(Chrome) 와 거의 동일한 헤더 세트.
        # sorisem(gnuboard) 의 일부 게시판은 권한 체크 시 봇 검출을 위해
        # Accept / Accept-Language 등 표준 헤더 유무를 본다. 누락 시 로그인은
        # 되어도 board.php 호출에서 "접근권한이 없습니다.(10000)" 응답이 오는
        # 사례가 발견되어, 가능한 한 브라우저 요청과 똑같이 맞춘다.
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,"
                "image/avif,image/webp,image/apng,*/*;q=0.8,"
                "application/signed-exchange;v=b3;q=0.7"
            ),
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Cache-Control": "max-age=0",
            "Sec-Ch-Ua": '"Chromium";v="120", "Not_A Brand";v="24"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "Connection": "keep-alive",
        })

    def authenticate(self, user_id: str, password: str) -> AuthResult:
        """
        소리샘 로그인 후 초록등대 동호회 회원 여부를 확인한다.

        Returns:
            AuthResult: 인증 결과
        """
        # 1단계: 소리샘 로그인
        login_result = self._login(user_id, password)
        if not login_result.is_success:
            return login_result

        # 2단계: 초록등대 동호회 회원 확인
        member_result = self._check_green_membership(user_id)
        return member_result

    def _login(self, user_id: str, password: str) -> AuthResult:
        """소리샘 사이트 로그인"""
        try:
            # 먼저 메인 페이지에 접속하여 세션 쿠키 획득
            self.session.get(SORISEM_BASE_URL, timeout=15)

            # 로그인 요청
            login_data = {
                "mb_id": user_id,
                "mb_password": password,
            }
            resp = self.session.post(
                LOGIN_URL,
                data=login_data,
                timeout=15,
                allow_redirects=True,
            )

            # 로그인 성공 여부 확인: 메인 페이지를 다시 요청하여 확인
            main_resp = self.session.get(SORISEM_BASE_URL, timeout=15)
            page_text = main_resp.text

            # 로그인 실패: 로그인 폼이 여전히 보이면 실패
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
        """초록등대 동호회 회원 목록에서 사용자 확인"""
        try:
            resp = self.session.get(GREEN_CLUB_MEMBERS_URL, timeout=15)
            soup = BeautifulSoup(resp.text, "lxml")

            # 회원 목록 테이블에서 아이디 검색
            # Gnuboard 동호회 회원 목록 페이지를 파싱
            page_text = resp.text.lower()

            # 회원 아이디가 페이지에 존재하는지 확인
            # 여러 페이지가 있을 수 있으므로 페이지네이션도 처리
            if self._find_member_in_page(soup, user_id):
                return AuthResult(
                    AuthResult.SUCCESS,
                    "초록등대 동호회 회원 인증에 성공했습니다."
                )

            # 페이지네이션 처리: 다음 페이지들도 확인
            page = 2
            while True:
                page_url = f"{GREEN_CLUB_MEMBERS_URL}&page={page}"
                resp = self.session.get(page_url, timeout=15)
                soup = BeautifulSoup(resp.text, "lxml")

                if self._find_member_in_page(soup, user_id):
                    return AuthResult(
                        AuthResult.SUCCESS,
                        "초록등대 동호회 회원 인증에 성공했습니다."
                    )

                # 더 이상 페이지가 없으면 종료
                if not self._has_next_page(soup, page):
                    break
                page += 1

            return AuthResult(
                AuthResult.NOT_MEMBER,
                "초록등대 동호회 회원이 아닙니다. 프로그램을 종료합니다."
            )

        except requests.exceptions.RequestException:
            return AuthResult(
                AuthResult.NETWORK_ERROR,
                "회원 목록을 확인할 수 없습니다. 인터넷 연결을 확인해 주세요."
            )

    def _find_member_in_page(self, soup: BeautifulSoup, user_id: str) -> bool:
        """페이지에서 회원 아이디를 찾는다."""
        # 방법 1: 테이블 셀에서 아이디 검색
        for td in soup.find_all("td"):
            cell_text = td.get_text(strip=True)
            if cell_text == user_id:
                return True

        # 방법 2: 링크 텍스트에서 아이디 검색
        for a_tag in soup.find_all("a"):
            link_text = a_tag.get_text(strip=True)
            if link_text == user_id:
                return True

        # 방법 3: 전체 텍스트에서 정확한 아이디 매칭
        # (단어 경계로 구분하여 부분 매칭 방지)
        import re
        full_text = soup.get_text()
        pattern = rf'\b{re.escape(user_id)}\b'
        if re.search(pattern, full_text):
            return True

        return False

    def _has_next_page(self, soup: BeautifulSoup, current_page: int) -> bool:
        """다음 페이지가 존재하는지 확인"""
        next_page = current_page + 1
        # 페이지 링크에서 다음 페이지 번호 검색
        for a_tag in soup.find_all("a", href=True):
            if f"page={next_page}" in a_tag["href"]:
                return True
        return False

    def logout(self) -> None:
        """소리샘 로그아웃"""
        try:
            from config import LOGOUT_URL
            self.session.get(LOGOUT_URL, timeout=10)
        except Exception:
            pass
        finally:
            self.session.cookies.clear()
