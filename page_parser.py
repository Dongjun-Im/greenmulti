"""소리샘 페이지 HTML 파싱 모듈"""
import re

from bs4 import BeautifulSoup

from config import SORISEM_BASE_URL


def _unwrap_js_url(href: str) -> str:
    """gnuboard 스타일의 javascript:delete_comment('/bbs/...?token=...')
    같은 래퍼에서 실제 URL만 추출한다. javascript: 가 아니면 그대로 반환.
    따옴표 안의 문자열이 여러 개일 경우 (예: confirm 메시지와 URL이 함께 있는
    형태) URL처럼 보이는 것을 우선 선택한다."""
    if not href:
        return ""
    h = href.strip()
    if not h.lower().startswith("javascript:"):
        return h
    # javascript:funcname('A', 'B', ...) 의 모든 quoted 문자열 후보 수집
    candidates = re.findall(r"""['"]([^'"]+)['"]""", h)
    if not candidates:
        return ""
    # URL로 보이는 후보(절대/상대/풀URL, .php 포함) 우선
    for c in candidates:
        cs = c.strip()
        if cs.startswith(("/", "./", "../", "http://", "https://")):
            return cs
        if ".php" in cs:
            return cs
    return candidates[0].strip()


def _is_real_delete_url(url: str) -> bool:
    """URL이 실제 삭제 엔드포인트를 가리키는지 확인.
    작성/수정 공용 엔드포인트(write_comment_update.php, write.php)를 잘못 잡아
    '댓글을 입력하여 주십시오' 같은 오류를 받는 문제를 방지한다."""
    if not url:
        return False
    u = url.lower()
    # 댓글 전용 삭제 엔드포인트 또는 일반 delete.php 만 인정
    return ("delete_comment" in u) or ("delete.php" in u)


def extract_post_author_id(html: str) -> str:
    """게시글 본문 페이지 HTML 에서 작성자의 mb_id(소리샘 로그인 아이디)를 추출.

    gnuboard 계열 사이트는 작성자 영역(.bo_v_info / .view_header 등) 안에 프로필
    링크 `.../profile.php?mb_id=XXX` 또는 `javascript:winprofile('XXX')` 형태로
    작성자 아이디를 노출한다. 이걸 바탕으로 현재 로그인한 사용자와 일치하는지
    클라이언트 측에서 검증하는 데 사용된다.

    추출 실패 시 빈 문자열. 호출자는 빈 값일 때 안전하게 거부(본인 아님으로 간주)
    하는 쪽이 기본 정책.
    """
    if not html:
        return ""
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        return ""

    # 게시글 상세 페이지에서 작성자 정보가 위치하는 후보 영역.
    author_region_selectors = [
        "#bo_v_info", ".bo_v_info",
        ".view_header", ".post_info",
        ".view_title", ".bo_v_tit",
        ".board_view_info",
    ]

    for sel in author_region_selectors:
        for el in soup.select(sel):
            for a in el.find_all("a", href=True):
                href = a["href"]
                # 1) URL 파라미터: ?mb_id=X, &mb_id=X
                m = re.search(r"[?&]mb_id=([^&'\"<>\s]+)", href)
                if m:
                    return m.group(1).strip()
                # 2) javascript 프로필 팝업: winprofile('X') / member_info('X')
                m = re.search(
                    r"(?:winprofile|member_info|mb_info)\s*\(\s*['\"]([^'\"]+)['\"]",
                    href,
                )
                if m:
                    return m.group(1).strip()
    return ""


def _extract_display_name(el) -> str:
    """HTML 요소에서 작성자 표시 이름(닉네임/이름)을 추출한다.
    아이디 대신 닉네임 우선."""
    if not el:
        return ""

    # 1. 이미지 alt 속성 (프로필 이미지에 닉네임이 있는 경우)
    img = el.find("img")
    if img:
        alt = img.get("alt", "").strip()
        if alt and 1 < len(alt) < 30 and not alt.lower().startswith(
            ("profile", "image", "icon", "photo", "avatar")
        ):
            return alt

    # 2. 닉네임 전용 클래스
    for sel in [".sv_name", ".if_name", ".mb_nick", ".nick",
                ".nickname", ".display_name", ".name"]:
        sub = el.select_one(sel)
        if sub:
            txt = sub.get_text(strip=True)
            if txt and 1 < len(txt) < 30:
                return txt

    # 3. <a> 태그의 title 속성 또는 텍스트
    a = el.find("a")
    if a:
        title = a.get("title", "").strip()
        if title and 1 < len(title) < 30:
            return title
        # sv_member 안의 텍스트 → 보통 닉네임
        txt = a.get_text(strip=True)
        if txt and 1 < len(txt) < 30:
            return txt

    # 4. 전체 텍스트 정제
    text = el.get_text(" ", strip=True)
    # "홍길동 (user123)" 형식: 괄호 밖 추출
    m = re.match(r'^([^()]+?)\s*\([^)]+\)\s*$', text)
    if m:
        return m.group(1).strip()
    # 한글 이름이 포함된 경우 한글 블록 우선
    kor = re.search(r'[\uAC00-\uD7A3][\uAC00-\uD7A3\s]*', text)
    if kor:
        cand = kor.group(0).strip()
        if 1 < len(cand) < 30:
            return cand
    return text.strip()


class SubMenuItem:
    """하위 메뉴 항목"""

    def __init__(self, name: str, url: str):
        self.name = name
        self.url = url

    @property
    def full_url(self) -> str:
        if self.url.startswith("http"):
            return self.url
        return f"{SORISEM_BASE_URL}{self.url}"

    @property
    def display_text(self) -> str:
        return self.name


class PostItem:
    """게시글 항목"""

    def __init__(self, number: str, title: str, author: str, date: str, url: str,
                 comment_count: int = 0):
        self.number = number
        self.title = title
        self.author = author
        self.date = date
        self.url = url
        self.comment_count = comment_count

    @property
    def full_url(self) -> str:
        if self.url.startswith("http"):
            return self.url
        return f"{SORISEM_BASE_URL}{self.url}"

    @property
    def display_text(self) -> str:
        """목록상자에 표시할 텍스트 (댓글 수 포함)"""
        parts = []
        if self.number:
            parts.append(self.number)
        parts.append(self.title)
        if self.comment_count > 0:
            parts.append(f"[댓글 {self.comment_count}]")
        if self.author:
            parts.append(f"- {self.author}")
        if self.date:
            parts.append(f"({self.date})")
        return " ".join(parts)


class CommentItem:
    """댓글 항목"""

    def __init__(self, author: str, date: str, body: str,
                 comment_id: str = "", edit_url: str = "", delete_url: str = ""):
        self.author = author
        self.date = date
        self.body = body
        self.comment_id = comment_id
        self.edit_url = edit_url
        self.delete_url = delete_url

    @property
    def display_text(self) -> str:
        return f"{self.author}: {self.body}"


class PostContent:
    """게시글 본문"""

    def __init__(self, title: str, author: str, date: str, body: str,
                 files: list[dict] | None = None,
                 comments: list[CommentItem] | None = None,
                 comment_write_url: str = "",
                 bo_table: str = "", wr_id: str = "",
                 prev_url: str = "", next_url: str = "",
                 edit_url: str = "", delete_url: str = "",
                 reply_url: str = ""):
        self.title = title
        self.author = author
        self.date = date
        self.body = body
        self.files = files or []
        self.comments = comments or []
        self.comment_write_url = comment_write_url
        self.bo_table = bo_table
        self.wr_id = wr_id
        self.prev_url = prev_url
        self.next_url = next_url
        self.edit_url = edit_url      # 게시물 수정 URL
        self.delete_url = delete_url  # 게시물 삭제 URL (토큰 포함)
        self.reply_url = reply_url    # 게시물 답변 URL


def _is_pagination_link(href: str, text: str) -> bool:
    """페이지네이션 링크인지 판별한다."""
    # page=만 있고 wr_id가 없는 링크만 페이지네이션
    # (게시글 링크에도 page=가 포함될 수 있으므로 wr_id 유무로 구분)
    if "page=" in href and "wr_id" not in href and "bo_table" not in href:
        return True
    # 숫자만 있는 텍스트 (페이지 번호)
    if re.match(r"^\d+$", text) and "wr_id" not in href:
        return True
    if text in ("이전", "다음", "처음", "끝", "prev", "next", "first", "last",
                "◀", "▶", "«", "»", "‹", "›"):
        return True
    return False


def _is_noise_link(href: str, text: str) -> bool:
    """무의미한 링크인지 판별한다."""
    if href in ("#", "javascript:void(0)", "javascript:;", ""):
        return True
    if "login" in href or "logout" in href or "register" in href:
        return True
    if "memo" in href and "popup" in href:
        return True
    if len(text) < 2 or len(text) > 50:
        return True

    # 스킵 네비게이션, 유틸리티, 하단 링크 제외
    noise_keywords = [
        "본문으로", "바로가기 이동", "본문 바로", "skip",
        "개인정보", "처리방침", "이용약관", "저작권",
        "쪽지", "메일보내기",
        "copyright", "privacy", "terms",
        "top", "위로", "맨위", "상단으로", "상단", "맨위로",
        "동사무소", "로그아웃",
    ]

    text_lower = text.lower()

    # 정확히 일치하는 짧은 노이즈 텍스트
    noise_exact = [
        "메일", "쪽지", "홈", "home", "돌아가기",
        "글쓰기", "게시판관리", "멀티업로드",
        "img", "관리자", "철머",
    ]
    if text_lower.strip() in noise_exact:
        return True

    for keyword in noise_keywords:
        if keyword in text_lower:
            return True

    # href가 메일/쪽지 관련인 경우
    noise_href_keywords = [
        "memo.php", "formmail", "mailto:",
        "member_confirm", "password",
    ]
    href_lower = href.lower()
    for keyword in noise_href_keywords:
        if keyword in href_lower:
            return True

    return False


def parse_sub_menus(html: str, base_url: str = "") -> list[SubMenuItem]:
    """
    페이지에서 하위 메뉴(서브 링크)를 추출한다.
    페이지네이션, 게시글 링크 등은 제외한다.

    base_url:
        현재 보고 있는 페이지의 URL. 전달되면 이 URL 이 가리키는 클럽/범위와
        다른 영역으로 향하는 링크를 걸러낸다. 예: base_url 이
        `/plugin/ar.club/?cl=green4` 이면 `cl=green4` 외의 `cl=다른값`을 가진
        링크는 최상위 내비게이션으로 간주하고 제외한다. 이를 통해 클럽 하위
        페이지에서 "일반 동호회", "초록 등대" 같은 상위 카테고리가 섞여
        표시되는 문제를 막는다.
    """
    soup = BeautifulSoup(html, "lxml")
    items = []
    seen_urls = set()

    # base_url 에서 현재 활성 클럽 코드(cl=X) 를 추출. 이후 필터에 사용.
    current_cl = ""
    if base_url:
        m = re.search(r"[?&]cl=([^&#]+)", base_url)
        if m:
            current_cl = m.group(1).strip().lower()

    def _is_out_of_scope(href: str) -> bool:
        """현재 클럽 컨텍스트에서 벗어난 상위 내비게이션인지 판정.

        소리샘은 `?mo=X&cl=Y` 형식으로 "Y 클럽의 X 섹션" 링크를 표현한다.
        따라서 부모 클럽을 가리키는 cl= 여도 mo= 가 함께 있으면 "부모의 특정
        섹션"이므로 breadcrumb이 아니라 실제 하위 메뉴 항목이다.

        판정 흐름:
          - cl= 없음 → 판단 불가, 통과
          - cl=<값> == current_cl → 같은 클럽, 통과
          - 공통 접두사 < 3자 → 무관한 카테고리, 제외
          - current_cl 이 cl_val 의 접두사 (부모 클럽) AND mo= 없음 → breadcrumb 제외
          - current_cl 이 cl_val 의 접두사 AND mo= 있음 → 부모의 섹션 링크, 통과
          - 그 외 (형제/자식 관계) → 통과
        """
        if not current_cl:
            return False
        href_lower = href.lower() if href else ""
        m2 = re.search(r"[?&]cl=([^&#]+)", href_lower)
        if not m2:
            return False
        cl_val = m2.group(1).strip().lower()
        if cl_val == current_cl:
            return False
        # 공통 접두사 길이 계산
        common_len = 0
        for c1, c2 in zip(cl_val, current_cl):
            if c1 == c2:
                common_len += 1
            else:
                break
        if common_len < 3:
            return True  # 무관한 클럽 (예: green → circle)
        has_mo = bool(re.search(r"[?&]mo=", href_lower))
        if current_cl.startswith(cl_val) and not has_mo:
            return True  # 부모 breadcrumb (예: green4 → green 단독 링크)
        return False  # 형제/자식/부모의 섹션(mo= 포함) → 통과

    def _collect_links(link_elements):
        for link in link_elements:
            href = link.get("href", "").strip()
            text = link.get_text(strip=True)

            if _is_noise_link(href, text):
                continue
            if _is_pagination_link(href, text):
                continue
            if "wr_id=" in href:
                continue
            # 현재 클럽과 무관한 다른 클럽/카테고리 링크 차단
            if _is_out_of_scope(href):
                continue

            # 소리샘 URL은 상대 경로로 변환, 외부 링크는 그대로 유지
            if href.startswith("http") and SORISEM_BASE_URL in href:
                href = href.replace(SORISEM_BASE_URL, "")

            if href not in seen_urls:
                seen_urls.add(href)
                items.append(SubMenuItem(text, href))

    # 1·2단계: 네비게이션 + 본문 영역 선택자에서 링크 추출. 두 단계를 병합해
    # 한쪽이 소수의 항목만 가진 경우에도 다른 쪽 항목을 놓치지 않는다. 예전에
    # 1단계에서 1~2 항목이 잡히면 2단계가 스킵되어 자식 클럽 목록을 통째로
    # 놓치는 문제를 방지한다.
    primary_selectors = [
        # nav
        "#c_menu a",
        "#c_aside a",
        ".side_menu a",
        "#aside a",
        ".snb a",
        ".lnb a",
        ".sub_menu a",
        ".category a",
        ".menu_list a",
        # content
        "#bo_list a",
        ".board_list a",
        "#content a",
        ".content a",
        "main a",
        "#container a",
        # club listings
        ".club_list a",
        ".art_club a",
        ".ar_club a",
        ".art_club_list a",
    ]

    for selector in primary_selectors:
        _collect_links(soup.select(selector))

    # 3단계를 항상 함께 실행 — 1·2단계가 일부만 포착할 때 누락을 방지.
    # header / footer / nav / 전역 메뉴 컨테이너를 제거한 사본을 따로 만들어
    # 전체 페이지를 스캔한다. seen_urls 기반 dedup 이 중복을 차단.
    soup_rest = BeautifulSoup(html, "lxml")
    for tag in soup_rest.find_all(["header", "footer", "script", "style", "nav"]):
        tag.decompose()
    for sel in (
        "#gnb", ".gnb", "#tnb", ".tnb", "#lnb", ".lnb", "#hd", ".hd",
        "#top", ".top", "#top_menu", ".top_menu",
        ".breadcrumb", ".breadcrumbs", ".crumb", ".location",
        "#globalNav", ".global-nav", ".global_nav",
    ):
        for node in soup_rest.select(sel):
            node.decompose()
    _collect_links(soup_rest.find_all("a", href=True))

    return items


def parse_board_list(html: str) -> list[PostItem]:
    """게시판 목록 페이지를 파싱하여 게시글 목록을 반환한다."""
    soup = BeautifulSoup(html, "lxml")
    posts = []

    # 방법 0: Gnuboard5 전용 (td_subject 셀렉터)
    for td in soup.select("td.td_subject, .td_subject"):
        link = td.select_one("a[href*='wr_id']")
        if not link:
            link = td.find("a", href=True)
        if not link:
            continue
        title = link.get_text(strip=True)
        href = link.get("href", "")
        if not title or len(title) < 2:
            continue

        # 댓글 수
        comment_count = 0
        cmt_el = td.select_one(".cnt_cmt, .comment_count")
        if cmt_el:
            cmt_text = cmt_el.get_text(strip=True)
            cmt_match = re.search(r'\d+', cmt_text)
            if cmt_match:
                comment_count = int(cmt_match.group())
            title = title.replace(cmt_text, "").strip()
        title = re.sub(r'\s*댓글\s*\d*\s*개\s*$', '', title).strip()

        # 같은 행에서 번호, 작성자, 날짜 추출
        tr = td.find_parent("tr")
        number = ""
        author = ""
        date = ""
        if tr:
            num_td = tr.select_one("td.td_num, .td_num")
            if num_td:
                number = num_td.get_text(strip=True)
            name_td = tr.select_one("td.td_name, .td_name, td.sv_member")
            if name_td:
                author = _extract_display_name(name_td)
            date_td = tr.select_one("td.td_date, .td_date, td.td_datetime")
            if date_td:
                date = date_td.get_text(strip=True)

        posts.append(PostItem(number, title, author, date, href, comment_count))

    if posts:
        return posts

    # 방법 1: 테이블 기반 게시판 (일반 스킨)
    for tr in soup.select("tbody tr, table tr"):
        cells = tr.find_all("td")
        if len(cells) < 2:
            continue

        # 제목과 링크 찾기
        title_cell = None
        title_link = None
        for cell in cells:
            for link in cell.find_all("a", href=True):
                href = link.get("href", "")
                # 게시글 링크: wr_id 또는 bo_table이 있는 링크
                if "wr_id" in href:
                    link_text = link.get_text(strip=True)
                    if len(link_text) > 1:
                        title_cell = cell
                        title_link = link
                        break
            if title_link:
                break

        if not title_link:
            continue

        title = title_link.get_text(strip=True)
        url = title_link.get("href", "")

        # 댓글 수: 제목 옆 (N) 또는 .cnt_cmt 등
        comment_count = 0
        if title_cell:
            cmt_el = title_cell.select_one(".cnt_cmt, .comment_count, .cmt_count")
            if cmt_el:
                cmt_text = cmt_el.get_text(strip=True)
                cmt_match = re.search(r'\d+', cmt_text)
                if cmt_match:
                    comment_count = int(cmt_match.group())
                # 댓글 수 텍스트를 제목에서 제거
                title = title.replace(cmt_text, "").strip()
        # 제목에서 (N) 패턴으로 댓글 수 추출
        if comment_count == 0:
            cmt_match = re.search(r'\s*\((\d+)\)\s*$', title)
            if cmt_match:
                comment_count = int(cmt_match.group(1))
                title = title[:cmt_match.start()].strip()

        # 제목에서 "댓글N개", "댓글 N개", "댓글개" 등 잔여 텍스트 제거
        title = re.sub(r'\s*댓글\s*\d*\s*개\s*$', '', title).strip()

        # 번호: 첫 번째 셀
        number = ""
        first_cell_text = cells[0].get_text(strip=True)
        if re.match(r"^\d+$", first_cell_text):
            number = first_cell_text

        # 작성자: 제목 셀 이후의 셀에서 찾기
        author = ""
        title_idx = cells.index(title_cell) if title_cell in cells else -1
        if title_idx >= 0:
            for i in range(title_idx + 1, len(cells)):
                cell_text = cells[i].get_text(strip=True)
                # 날짜가 아니고, 숫자만이 아니고, 짧은 텍스트면 작성자
                if (cell_text and
                    not re.match(r"^\d+$", cell_text) and
                    not re.match(r"\d{2,4}[-/.]\d{2}[-/.]\d{2}", cell_text) and
                    len(cell_text) < 20):
                    author = cell_text
                    break

        # 날짜
        date = ""
        for cell in cells:
            cell_text = cell.get_text(strip=True)
            if re.match(r"\d{2,4}[-/.]\d{2}([-/.]\d{2})?", cell_text):
                date = cell_text
                break

        posts.append(PostItem(number, title, author, date, url, comment_count))

    # 방법 2: 리스트 기반 게시판 (ul/li 스킨)
    if not posts:
        for li in soup.select("ul li"):
            link = li.find("a", href=True)
            if not link:
                continue

            href = link.get("href", "")
            title = link.get_text(strip=True)

            # 게시글 링크만 선택
            if "wr_id" not in href:
                continue
            if _is_pagination_link(href, title):
                continue
            if len(title) < 2:
                continue

            author = ""
            date = ""
            for span in li.find_all("span"):
                span_text = span.get_text(strip=True)
                if re.match(r"\d{2,4}[-/.]\d{2}", span_text):
                    date = span_text
                elif span_text and span_text != title and len(span_text) < 20:
                    author = span_text

            posts.append(PostItem("", title, author, date, href))

    # 방법 3: 페이지 전체에서 wr_id 링크 수집 (최후의 수단)
    if not posts:
        for a_tag in soup.find_all("a", href=True):
            href = a_tag.get("href", "")
            title = a_tag.get_text(strip=True)
            if "wr_id" in href and title and len(title) > 1:
                if not _is_pagination_link(href, title):
                    posts.append(PostItem("", title, "", "", href))

    return posts


def parse_post_content(html: str) -> PostContent | None:
    """게시글 본문 페이지를 파싱한다."""
    soup = BeautifulSoup(html, "lxml")

    # 제목: 여러 선택자 시도
    title = ""
    for selector in [".bo_v_tit", "#bo_v_title", ".view_title",
                     "h2.title", ".subject", "h1", "h2", "h3"]:
        el = soup.select_one(selector)
        if el:
            text = el.get_text(strip=True)
            if text and len(text) > 1:
                title = text
                break

    # 작성자 (닉네임/이름 우선, 아이디 최후순위)
    author = ""
    for selector in [
        ".bo_v_info .sv_member", ".sv_member",
        ".bo_v_info .if_name", ".if_name",
        ".bo_v_info .mb_nick", ".mb_nick",
        ".bo_v_info .writer", ".writer",
        ".bo_v_info .name", ".name",
        ".author", ".post_info .writer",
        ".write_info .name",
    ]:
        el = soup.select_one(selector)
        if el:
            text = _extract_display_name(el)
            if text and len(text) < 40:
                author = text
                break

    # 날짜 + 시간 (title 속성에 풀 datetime이 있으면 우선)
    date = ""
    for selector in [
        ".bo_v_info .sv_date", ".sv_date",
        ".bo_v_info .if_date", ".if_date",
        ".bo_v_info .date", ".date", ".datetime",
        ".bo_v_info time", "time",
        ".reg_date", ".post_date",
    ]:
        el = soup.select_one(selector)
        if el:
            title_attr = el.get("title", "").strip()
            dt_attr = el.get("datetime", "").strip()
            text = el.get_text(strip=True)
            candidates = [c for c in (title_attr, dt_attr, text) if c]
            if candidates:
                date = max(candidates, key=len)
                break

    # 정보 영역(.bo_v_info 등) 보완 탐색
    info_area = soup.select_one(
        "#bo_v_info, .bo_v_info, .view_info, .post_info, .post_meta, "
        ".bo_v_nb, .write_info, #write_info"
    )
    if info_area:
        info_text_all = info_area.get_text(" ", strip=True)

        # 1) 정보 영역 전체에서 이름 우선 추출
        if not author:
            candidate = _extract_display_name(info_area)
            # 후보가 너무 길거나 라벨 단어만이면 제외
            if (
                candidate
                and 1 < len(candidate) < 30
                and not re.search(r'작성|날짜|조회|IP|추천', candidate)
            ):
                author = candidate

        # 2) "작성자 XXX", "글쓴이 XXX", "이름 XXX" 패턴
        if not author:
            m = re.search(
                r'(?:작성자|글쓴이|이름)\s*[:：]?\s*([^\s|·,]{1,30})',
                info_text_all,
            )
            if m:
                author = m.group(1).strip()

        # 3) <strong>작성자</strong>XXX 패턴
        if not author:
            for strong in info_area.find_all("strong"):
                label = strong.get_text(strip=True).replace(":", "").strip()
                if label in ("작성자", "글쓴이", "이름"):
                    sibling_text = ""
                    for sib in strong.next_siblings:
                        t = (
                            sib.get_text(strip=True)
                            if hasattr(sib, "get_text")
                            else str(sib).strip()
                        )
                        if t:
                            sibling_text = t
                            break
                    if sibling_text:
                        author = sibling_text.strip()
                        break

        # 날짜 보완: 전체 텍스트에서 날짜 + 시간 함께 검색
        if not date or not re.search(r'\d{1,2}:\d{2}', date):
            full_date_match = re.search(
                r'(\d{2,4}[-./]\d{1,2}[-./]\d{1,2})\s*(\d{1,2}:\d{2}(?::\d{2})?)?',
                info_text_all,
            )
            if full_date_match:
                d = full_date_match.group(1)
                t = full_date_match.group(2) or ""
                if t:
                    date = f"{d} {t}"
                elif not date:
                    date = d

    # 여전히 시간이 없으면 페이지 전체에서 날짜 + 시간 조합 탐색
    if date and not re.search(r'\d{1,2}:\d{2}', date):
        page_text = soup.get_text(" ", strip=True)
        # 이미 추출한 날짜 문자열과 근접한 시간을 찾는다
        date_core = re.search(r'\d{2,4}[-./]\d{1,2}[-./]\d{1,2}', date)
        if date_core:
            core = date_core.group(0)
            idx = page_text.find(core)
            if idx >= 0:
                after = page_text[idx:idx + 60]
                m = re.search(r'\d{1,2}:\d{2}(?::\d{2})?', after)
                if m:
                    date = f"{core} {m.group(0)}"

    # 본문: 여러 선택자 시도
    body = ""
    for selector in ["#bo_v_con", ".bo_v_con", ".view_content",
                     "#writeContents", ".content_view", "#bo_v_atc"]:
        body_el = soup.select_one(selector)
        if body_el:
            for br in body_el.find_all("br"):
                br.replace_with("\n")
            for p in body_el.find_all("p"):
                p.insert_after("\n")
            body = body_el.get_text().strip()
            if body:
                break

    # 첨부파일: 명확한 파일 다운로드 영역에서만 추출
    files = []
    file_container_selectors = [
        ".bo_v_file", "#bo_v_file",
        ".file_list", ".view_file",
        ".file_wrap", "#file_wrap",
    ]
    file_container = None
    for selector in file_container_selectors:
        file_container = soup.select_one(selector)
        if file_container:
            break

    if file_container:
        for file_link in file_container.find_all("a", href=True):
            file_name = file_link.get_text(strip=True)
            file_url = file_link.get("href", "")
            if (file_name and file_url
                    and not _is_pagination_link(file_url, file_name)
                    and not _is_noise_link(file_url, file_name)
                    and len(file_name) > 1):
                # 파일명에서 용량 정보 제거
                # 괄호 포함: "(378byte)", "(21.6KB)"
                file_name = re.sub(
                    r'\s*\(\d+[\.\d]*\s*[BbKkMmGg][Bb]?[Yy]?[Tt]?[Ee]?[Ss]?\)\s*$',
                    '', file_name
                ).strip()
                # 괄호 없이: "378byte", "21.6k"
                file_name = re.sub(
                    r'\s+\d+[\.\d]*\s*[BbKkMmGg][Bb]?[Yy]?[Tt]?[Ee]?[Ss]?\s*$',
                    '', file_name
                ).strip()
                if file_name:
                    files.append({"name": file_name, "url": file_url})

    if not title and not body:
        return None

    # bo_table과 wr_id 추출 (댓글 작성 시 필요)
    bo_table = ""
    wr_id = ""
    # URL에서 추출
    for form in soup.find_all("form"):
        action = form.get("action", "")
        if "comment" in action or "bo_table" in action:
            bo_input = form.find("input", {"name": "bo_table"})
            wr_input = form.find("input", {"name": "wr_id"})
            if bo_input:
                bo_table = bo_input.get("value", "")
            if wr_input:
                wr_id = wr_input.get("value", "")
            break

    # hidden input에서도 시도
    if not bo_table:
        bo_input = soup.find("input", {"name": "bo_table"})
        if bo_input:
            bo_table = bo_input.get("value", "")
    if not wr_id:
        wr_input = soup.find("input", {"name": "wr_id"})
        if wr_input:
            wr_id = wr_input.get("value", "")

    # 페이지 URL에서 추출 (최후의 수단)
    if not bo_table or not wr_id:
        for a_tag in soup.find_all("a", href=True):
            href = a_tag.get("href", "")
            if "bo_table=" in href and "wr_id=" in href:
                bo_match = re.search(r'bo_table=([^&]+)', href)
                wr_match = re.search(r'wr_id=(\d+)', href)
                if bo_match and not bo_table:
                    bo_table = bo_match.group(1)
                if wr_match and not wr_id:
                    wr_id = wr_match.group(1)
                if bo_table and wr_id:
                    break

    # 댓글 파싱 (에러가 나도 게시물은 정상 반환)
    comments = []
    try:
        comments = _parse_comments(soup)
        if comments:
            comments = _filter_valid_comments(comments)
    except Exception:
        comments = []

    # 댓글 작성 URL
    comment_write_url = ""
    comment_form = soup.select_one(
        "#fwrite_comment, form[name='fwrite'], #comment_form, "
        "form[action*='comment_update']"
    )
    if comment_form:
        comment_write_url = comment_form.get("action", "")

    # 이전글/다음글 URL 파싱
    prev_url = ""
    next_url = ""
    for a_tag in soup.find_all("a", href=True):
        a_text = a_tag.get_text(strip=True)
        href = a_tag.get("href", "")
        if not href or "wr_id" not in href:
            continue
        if "이전" in a_text or "prev" in a_text.lower():
            if not prev_url:
                prev_url = href
        elif "다음" in a_text or "next" in a_text.lower():
            if not next_url:
                next_url = href

    # 게시물 수정/삭제/답변 URL 추출
    edit_url = ""
    delete_url = ""
    reply_url = ""
    for a_tag in soup.find_all("a", href=True):
        href = a_tag.get("href", "")
        a_text = a_tag.get_text(strip=True)
        if "write.php" in href and "w=u" in href and not edit_url:
            edit_url = href
        elif "delete.php" in href and not delete_url:
            delete_url = href
        elif "write.php" in href and "w=r" in href and not reply_url:
            reply_url = href

    return PostContent(title, author, date, body, files, comments,
                       comment_write_url, bo_table, wr_id,
                       prev_url, next_url,
                       edit_url, delete_url, reply_url)


def _parse_comments(soup: BeautifulSoup) -> list[CommentItem]:
    """게시글 페이지에서 댓글 목록을 파싱한다."""
    comments = []

    # 소리샘 Gnuboard5 댓글 구조:
    # <section id="bo_vc">
    #   <article id="c_NNNN">
    #     <p>작성자님 날짜</p>
    #     <p>댓글 본문</p>
    #     <textarea style="display:none">댓글 본문 (수정용 복사)</textarea>
    #     <footer>수정/삭제 버튼</footer>
    #   </article>
    # </section>

    bo_vc = soup.select_one("#bo_vc, section#bo_vc, .bo_vc")
    if not bo_vc:
        return []

    for article in bo_vc.select("article"):
        # <textarea> 내용 백업 (수정용 사본) 후 제거 - 본문 누락 방지
        textarea_body = ""
        for ta in article.select("textarea"):
            textarea_body = ta.get_text("\n", strip=True)
            ta.decompose()
        # <footer>: 수정/삭제 URL 추출 후 제거
        edit_url = ""
        delete_url = ""
        footer = article.select_one("footer")
        if footer:
            for a in footer.find_all("a", href=True):
                h = a.get("href", "")
                t = a.get_text(strip=True)
                if "수정" in t or "edit_comment" in h:
                    edit_url = _unwrap_js_url(h)
                elif "삭제" in t or "delete_comment" in h:
                    delete_url = _unwrap_js_url(h)
            footer.decompose()
        # 본인 댓글은 수정/삭제 버튼이 <ul>이나 inline a에 있을 수 있어 추가 제거
        for el in article.select("ul.btn_confirm, ul.comment_btns, .btn_area, .cmt_btn"):
            for a in el.find_all("a", href=True):
                h = a.get("href", "")
                t = a.get_text(strip=True)
                if not edit_url and ("수정" in t or "edit_comment" in h):
                    edit_url = _unwrap_js_url(h)
                if not delete_url and ("삭제" in t or "delete_comment" in h):
                    delete_url = _unwrap_js_url(h)
            el.decompose()

        # <p> 태그들에서 정보 추출
        p_tags = article.find_all("p")

        author = ""
        date = ""
        body = ""

        if len(p_tags) >= 2:
            # 첫 번째 <p>: 작성자 + 날짜
            header_text = p_tags[0].get_text(strip=True)
            # "이름 (아이디)님 YY-MM-DD HH:MM" 패턴
            h_match = re.match(
                r'(.+?님?)\s+(\d{2,4}-\d{2}-\d{2}\s+\d{2}:\d{2})',
                header_text
            )
            if h_match:
                author = h_match.group(1).rstrip("님").strip()
                date = h_match.group(2).strip()
            else:
                author = header_text

            # 두 번째 <p>: 본문
            for br in p_tags[1].find_all("br"):
                br.replace_with("\n")
            body = p_tags[1].get_text(strip=True)

        elif len(p_tags) == 1:
            # <p> 하나만 있으면 전체 텍스트에서 분리
            full = p_tags[0].get_text(strip=True)
            h_match = re.match(
                r'(.+?님?)\s+(\d{2,4}-\d{2}-\d{2}\s+\d{2}:\d{2})\s*(.*)',
                full, re.DOTALL
            )
            if h_match:
                author = h_match.group(1).rstrip("님").strip()
                date = h_match.group(2).strip()
                body = h_match.group(3).strip()
            else:
                body = full

        # body가 비어있으면 textarea 백업 사용 (본인 댓글 편집용 사본)
        if not body and textarea_body:
            body = textarea_body

        if not body and not author:
            continue

        comment_id = article.get("id", "")

        comments.append(CommentItem(author, date, body, comment_id,
                                    edit_url, delete_url))

    return _filter_valid_comments(comments)


def _extract_single_comment(el) -> CommentItem | None:
    """단일 HTML 요소에서 댓글 정보를 추출한다."""
    author = ""
    date = ""
    body = ""

    # 작성자
    for sel in [".sv_member", ".comment_name", ".cmt_name",
                ".name", ".writer", ".nickname"]:
        a_el = el.select_one(sel)
        if a_el:
            author = a_el.get_text(strip=True)
            if author.endswith("님"):
                author = author[:-1].strip()
            if author:
                break

    # 날짜
    for sel in [".sv_date", ".comment_date", ".cmt_date", ".date", "time"]:
        d_el = el.select_one(sel)
        if d_el:
            date = d_el.get_text(strip=True)
            if date:
                break

    # 본문: 댓글 전용 셀렉터만 사용 (.content, .text는 게시글 본문과 혼동)
    for sel in [".comment_content", ".cmt_text", ".cmt_contents",
                ".usermemo", ".bo_vc_con", ".comment_text"]:
        b_el = el.select_one(sel)
        if b_el:
            for br in b_el.find_all("br"):
                br.replace_with(" ")
            body = b_el.get_text(strip=True)
            if body:
                break

    # 폴백: 전체 텍스트에서 작성자/날짜/버튼 제거하여 본문 추출
    if not body:
        # 작성자/날짜 요소를 제거한 복사본에서 텍스트 추출
        import copy
        try:
            el_copy = copy.copy(el)
            for remove_sel in [".sv_member", ".sv_date", ".comment_head",
                               ".comment_name", ".cmt_name", ".date"]:
                for rem in el_copy.select(remove_sel):
                    rem.decompose()
            # 버튼 영역도 제거
            for rem in el_copy.select("a"):
                t = rem.get_text(strip=True)
                if t in ("수정", "삭제", "답변", "신고"):
                    rem.decompose()
            body = el_copy.get_text(strip=True)
        except Exception:
            full = el.get_text(strip=True)
            if author:
                full = full.replace(author + "님", "", 1)
                full = full.replace(author, "", 1)
            if date:
                full = full.replace(date, "", 1)
            for noise in ["수정", "삭제", "답변", "신고", "님"]:
                full = full.replace(noise, "")
            body = full.strip()

        if body == author or body == date:
            body = ""

    if not body and not author:
        return None

    # 수정/삭제 URL
    comment_id = el.get("id", "")
    edit_url = ""
    delete_url = ""
    for a in el.find_all("a", href=True):
        h = a.get("href", "")
        t = a.get_text(strip=True)
        if "수정" in t or "edit" in h.lower():
            edit_url = _unwrap_js_url(h)
        elif "삭제" in t or "delete" in h.lower():
            delete_url = _unwrap_js_url(h)

    return CommentItem(author, date, body, comment_id, edit_url, delete_url)

    # 1단계: 기타 댓글 영역
    comment_selectors = [
        "#bo_vc li",
        ".bo_vc li",
        "#bo_vc > div",
        ".bo_vc > div",
        ".bo_vc_wrap",
        "#comment_list li",
        ".comment_list li",
        ".cmt_list li",
    ]

    comment_elements = []
    for selector in comment_selectors:
        comment_elements = soup.select(selector)
        if comment_elements:
            break

    # 2단계: id/class에 comment/cmt/vc가 포함된 영역 내 하위 요소
    if not comment_elements:
        for container in soup.find_all(True, id=re.compile(r'comment|cmt|reply|bo_vc', re.I)):
            children = container.find_all(["li", "div", "article", "section"], recursive=False)
            if children:
                comment_elements = children
                break
        if not comment_elements:
            for container in soup.find_all(True, class_=re.compile(r'comment|cmt|reply|bo_vc', re.I)):
                children = container.find_all(["li", "div", "article", "section"], recursive=False)
                if children:
                    comment_elements = children
                    break

    return _parse_comment_elements(comment_elements)


def _filter_valid_comments(comments: list[CommentItem]) -> list[CommentItem]:
    """잘못 파싱된 댓글(이전글/다음글, 사이트 정보 등)을 제거한다."""
    # 실제 댓글 본문에 나타나기 어려운 사이트 정보/네비게이션 텍스트만 필터
    # (수정/삭제/답변/목록/글쓰기 등은 일상어라 본문에도 나올 수 있어 제외함)
    noise_patterns = [
        "이전글", "다음글", "이전 글", "다음 글",
        "TEL", "FAX", "e-mail", "sorisem",
        "서울특별시", "copyright",
        "개인정보 처리방침", "이용약관",
        "상단으로", "본문으로",
    ]

    valid = []
    for c in comments:
        text = f"{c.author} {c.body}".strip()
        if not text or len(text) < 2:
            continue
        # 노이즈 패턴이 포함되어 있으면 제외
        is_noise = False
        for pattern in noise_patterns:
            if pattern in text:
                is_noise = True
                break
        if not is_noise:
            valid.append(c)
    return valid


def _parse_comment_elements(elements) -> list[CommentItem]:
    """HTML 요소 목록에서 댓글을 추출한다."""
    comments = []

    for el in elements:
        # 작성자
        author = ""
        for sel in [".sv_member", ".comment_name", ".cmt_name",
                    ".name", ".writer", ".nickname",
                    "span.name", "a.member", "b", "strong"]:
            author_el = el.select_one(sel)
            if author_el:
                author = author_el.get_text(strip=True)
                if author:
                    break

        # 날짜
        date = ""
        for sel in [".sv_date", ".comment_date", ".cmt_date",
                    ".date", ".datetime", "time", "span.date"]:
            date_el = el.select_one(sel)
            if date_el:
                date = date_el.get_text(strip=True)
                if date:
                    break

        # 본문
        body = ""
        for sel in [".sv_wrap", ".comment_text", ".cmt_text",
                    ".cmt_contents", ".comment_content",
                    ".content", ".text", "p"]:
            body_el = el.select_one(sel)
            if body_el:
                body = body_el.get_text(strip=True)
                if body:
                    break

        # 본문 폴백: 요소 전체 텍스트에서 작성자/날짜/버튼 텍스트 제거
        if not body:
            full_text = el.get_text(strip=True)
            if author and author in full_text:
                full_text = full_text.replace(author, "", 1).strip()
            if date and date in full_text:
                full_text = full_text.replace(date, "", 1).strip()
            for noise in ["수정", "삭제", "답변", "신고"]:
                full_text = full_text.replace(noise, "").strip()
            body = full_text

        if not body and not author:
            continue

        # 댓글 ID
        comment_id = el.get("id", "")
        if not comment_id:
            id_match = re.search(r'comment_(\d+)|c_(\d+)|wr_id=(\d+)',
                                 str(el.get("id", "")) + " " + str(el))
            if id_match:
                comment_id = id_match.group(1) or id_match.group(2) or id_match.group(3)

        # 수정/삭제 URL
        edit_url = ""
        delete_url = ""
        for a_tag in el.find_all("a", href=True):
            href = a_tag.get("href", "")
            a_text = a_tag.get_text(strip=True)
            if "수정" in a_text or "edit" in href.lower():
                edit_url = _unwrap_js_url(href)
            elif "삭제" in a_text or "delete" in href.lower():
                delete_url = _unwrap_js_url(href)

        comments.append(CommentItem(author, date, body, comment_id,
                                    edit_url, delete_url))

    return comments
