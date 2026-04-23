"""쪽지(memo) API + 대화상자.

소리샘(gnuboard5) 표준 memo 플러그인 기반. 받은 쪽지함·보낸 쪽지함 조회,
개별 쪽지 읽기·삭제·작성·답장을 제공.

엔드포인트:
- GET  /bbs/memo.php?kind=recv|send  목록
- GET  /bbs/memo_view.php?me_id=N    개별 보기
- GET  /bbs/memo_form.php            작성 폼 (token 획득)
- POST /bbs/memo_form_update.php     전송
- GET  /bbs/memo_delete.php?me_id=N  삭제
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import urljoin

import requests
import wx
from bs4 import BeautifulSoup

from config import (
    SORISEM_BASE_URL,
    MEMO_LIST_URL, MEMO_VIEW_URL, MEMO_FORM_URL,
    MEMO_FORM_UPDATE_URL, MEMO_DELETE_URL, MEMO_LIST_UPDATE_URL,
    MEMO_CHECK_NEW_URL, MEMO_NOTIFY_INTERVAL_SEC,
)
from screen_reader import speak
from theme import apply_theme, make_font, load_font_size


# ── 데이터 클래스 ──

@dataclass
class MemoItem:
    """목록 한 줄."""
    me_id: str            # 쪽지 ID (URL 파라미터)
    counterpart: str      # 받은: 보낸이, 보낸: 받는이
    summary: str          # 본문 요약 또는 제목 (gnuboard5 는 제목이 없고 본문 앞 일부)
    date: str             # 전송일
    is_read: bool = True  # 받은 쪽지함에서만 의미 있음


@dataclass
class MemoContent:
    """개별 쪽지 내용."""
    me_id: str
    sender: str           # 보낸이 아이디/닉네임
    recipient: str        # 받는이
    date: str
    body: str
    kind: str = "recv"    # recv / send


# ── HTML 파싱 ──

def _parse_memo_list(html: str) -> list[MemoItem]:
    """ar.memo 플러그인 목록 HTML 을 MemoItem 리스트로.

    실제 HTML 구조 (소리샘 2026-04 기준):
    <div id="memo_list">
      <form id="fboardlist" action="./memo_list_update.php">
        <input name="memo-type" value="recv">
        <div class="tbl_head01 tbl_wrap">
          <table>
            <thead><tr><th>체크</th><th>보낸사람</th><th>내용</th><th>시간</th><th>읽음</th></tr></thead>
            <tbody>
              <tr>
                <td><input name="chk_me_id[]" value="12345"></td>
                <td>sender_id</td>
                <td><a href="./memo_view.php?me_id=12345">본문 요약...</a></td>
                <td>2026-04-23 10:30</td>
                <td>O / X</td>
              </tr>
              ... or ...
              <tr><td colspan="4" class="empty_table">자료가 없습니다.</td></tr>
            </tbody>
          </table>
        </div>
      </form>
    </div>
    """
    soup = BeautifulSoup(html, "lxml")
    items: list[MemoItem] = []

    # 메모 테이블 찾기 — 여러 셀렉터 순차 시도
    table = None
    for sel in [
        "#memo_list form .tbl_head01 table",
        "#memo_list .tbl_head01 table",
        "#memo_list table",
        ".tbl_head01 table",
        "form#fboardlist table",
    ]:
        table = soup.select_one(sel)
        if table:
            break
    if not table:
        # 최후: 페이지 어딘가의 첫 테이블
        table = soup.find("table")
    if not table:
        return items

    tbody = table.find("tbody")
    if not tbody:
        return items

    for tr in tbody.find_all("tr", recursive=False):
        # 빈 목록 표시 행 스킵
        if tr.find("td", class_="empty_table"):
            continue
        # td colspan > 1 으로 된 안내성 행 스킵 (안전장치)
        empty_td = tr.find("td", attrs={"colspan": True})
        if empty_td and "자료가 없습니다" in empty_td.get_text(strip=True):
            continue

        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) < 4:
            continue

        # me_id 추출 — 체크박스 우선, 없으면 보기 링크
        me_id = ""
        chk = tr.find("input", {"name": re.compile(r"chk_me_id")})
        if chk and chk.get("value"):
            me_id = chk.get("value")
        if not me_id:
            view_a = tr.find("a", href=re.compile(r"memo_view\.php"))
            if view_a:
                m = re.search(r"me_id=(\d+)", view_a.get("href", ""))
                if m:
                    me_id = m.group(1)
        if not me_id:
            continue

        def cell_text(idx):
            if idx >= len(cells):
                return ""
            return cells[idx].get_text(" ", strip=True)

        # 컬럼 순서: [0]체크박스 [1]보낸사람 [2]내용 [3]시간 [4]읽음
        counterpart = cell_text(1)
        summary = cell_text(2)
        date = cell_text(3)
        read_cell = cell_text(4)

        # 안 읽음 판단 — ar.memo 는 읽음 컬럼에 "아직 읽지 않음" / 읽은 시각 표시.
        # gnuboard5 표준 테마는 "O"/"X" 또는 envelope 아이콘.
        is_read = True
        row_html = str(cells[4]) if len(cells) > 4 else ""
        if re.search(r"읽지\s*않|안\s*읽|미열람|미확인|미읽음", read_cell):
            is_read = False
        elif read_cell.strip() in ("X", "x", "", "-"):
            is_read = False
        elif "envelope-o" in row_html and "envelope-open" not in row_html:
            is_read = False

        items.append(MemoItem(
            me_id=me_id,
            counterpart=counterpart or "(알 수 없음)",
            summary=summary or "(내용 없음)",
            date=date or "",
            is_read=is_read,
        ))
    return items


def _parse_memo_content(html: str, kind: str = "recv") -> Optional[MemoContent]:
    """개별 쪽지 HTML 을 MemoContent 로.

    ar.memo 플러그인의 실제 구조 (2026-04 확인):
    <article id="memo_view_contents">
        <header><h1>메모 내용</h1></header>
        <ul id="memo_view_ul">
            <li class="memo_view_li">
                <span class="memo_view_subj">받는사람</span>
                <strong>anycall (임동준)</strong>
            </li>
            <li class="memo_view_li">
                <span class="memo_view_subj">보낸시간</span>
                <strong>26-04-23 13:47</strong>
                (읽지않음)
            </li>
        </ul>
        <p>본문 텍스트</p>
    </article>
    <div class="_win_btn">
        <a>다음 메모</a> <a>답장</a> <a>삭제</a> <a>목록보기</a>
    </div>
    """
    soup = BeautifulSoup(html, "lxml")

    sender = ""
    recipient = ""
    date = ""
    body = ""

    # ── 우선 순위 0: ar.memo 전용 구조 (#memo_view_contents) ──
    article = soup.select_one("#memo_view_contents")
    if article:
        for li in article.select("li.memo_view_li"):
            subj_el = li.find("span", class_="memo_view_subj")
            if not subj_el:
                continue
            label = subj_el.get_text(" ", strip=True)
            strong = li.find("strong")
            if strong:
                value_raw = strong.get_text(" ", strip=True)
            else:
                value_raw = li.get_text(" ", strip=True)
                value_raw = value_raw.replace(label, "", 1).strip()
            if not value_raw:
                continue
            if "보낸사람" in label or "보낸이" in label:
                sender = value_raw
            elif "받는사람" in label or "받는이" in label:
                recipient = value_raw
            elif "보낸시간" in label or "시간" in label or "작성" in label or "날짜" in label:
                # "26-04-23 13:47 (읽지않음)" 처럼 뒤에 부가 텍스트 붙을 수 있어 날짜만 추출
                m = re.match(r'([\d\-]+(?:\s+[\d:]+)?)', value_raw)
                date = m.group(1).strip() if m else value_raw.split("(")[0].strip()
        # 본문 — <p> 태그들
        body_parts = []
        for p in article.find_all("p"):
            text_p = p.get_text("\n", strip=True)
            if text_p:
                body_parts.append(text_p)
        if body_parts:
            body = "\n\n".join(body_parts)
        if body or sender or recipient:
            return MemoContent(
                me_id="", sender=sender, recipient=recipient,
                date=date, body=body or "(본문 없음)", kind=kind,
            )

    # ── 폴백 경로 계속 ──
    container = (soup.select_one("#memo_view")
                 or soup.select_one(".new_win.mbskin")
                 or soup.find(id=re.compile(r"memo", re.I))
                 or soup)

    # 우선 순위 1: <th>/<td> 페어 (가장 정확)
    for tr in container.find_all("tr"):
        th = tr.find("th")
        td = tr.find("td")
        if not th or not td:
            continue
        label = th.get_text(" ", strip=True)
        # td 내부의 <br> 를 줄바꿈으로 변환
        for br in td.find_all("br"):
            br.replace_with("\n")
        value = td.get_text("\n", strip=True)
        if not value:
            continue
        if "보낸" in label and not sender:
            sender = value
        elif "받는" in label and not recipient:
            recipient = value
        elif any(k in label for k in ["시간", "작성", "날짜", "일시", "보낸시간", "보낸 시간"]) and not date:
            date = value
        elif any(k in label for k in ["내용", "본문", "메시지", "메모"]) and not body:
            body = value

    # 우선 순위 2: dl/dt/dd 페어 (일부 테마)
    if not body:
        for dl in container.find_all("dl"):
            dt_list = dl.find_all("dt")
            dd_list = dl.find_all("dd")
            for dt, dd in zip(dt_list, dd_list):
                label = dt.get_text(" ", strip=True)
                for br in dd.find_all("br"):
                    br.replace_with("\n")
                value = dd.get_text("\n", strip=True)
                if not value:
                    continue
                if "보낸" in label and not sender:
                    sender = value
                elif "받는" in label and not recipient:
                    recipient = value
                elif any(k in label for k in ["시간", "작성", "날짜", "일시"]) and not date:
                    date = value
                elif any(k in label for k in ["내용", "본문", "메시지"]) and not body:
                    body = value

    # 우선 순위 3: 컨테이너에서 네비·제목·버튼을 제거한 뒤 텍스트 추출
    # (th/td 구조가 전혀 없는 예외적 테마용 폴백)
    if not body and not sender:
        import copy as _copy
        clean = _copy.copy(container)
        # 네비게이션 제거
        for sel in [".win_ul", "ul.win_ul", ".memo_nav", ".nav-memo"]:
            for el in clean.select(sel):
                el.decompose()
        # 제목 제거
        for tag in clean.find_all(["h1", "h2", "h3", "h4"]):
            tag.decompose()
        # 액션 버튼 영역 제거
        for sel in [".btn_confirm", ".btn_bo_user", ".btn_bo_adm", ".btn_wrap",
                    ".bo_fx", ".memo_btn", "._win_btn", "form"]:
            for el in clean.select(sel):
                el.decompose()
        # '답장', '삭제', '목록보기' 텍스트 링크 제거
        for a in clean.find_all("a"):
            t = a.get_text(strip=True)
            if t in ("답장", "삭제", "목록보기", "목록", "답변"):
                a.decompose()
        # <br> 줄바꿈
        for br in clean.find_all("br"):
            br.replace_with("\n")
        text = clean.get_text("\n", strip=True)
        # 메타 정보 정규식 추출
        m = re.search(r"보낸\s*(?:사람|이|분|사용자)\s*[:：]?\s*(\S+)", text)
        if m:
            sender = m.group(1)
        m = re.search(r"받는\s*(?:사람|이|분|사용자)\s*[:：]?\s*(\S+)", text)
        if m:
            recipient = m.group(1)
        m = re.search(r"(\d{4}-\d{2}-\d{2}[\s\d:]*)", text)
        if m:
            date = m.group(1).strip()
        # 메타 라인들을 본문에서 제거하고 나머지를 본문으로
        body_lines = []
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            # 메타 라벨 행은 건너뛰기
            if re.match(r"^(보낸|받는)\s*(사람|이|분|사용자)", line):
                continue
            if re.match(r"^(시간|날짜|작성일|일시)", line):
                continue
            if line in (sender, recipient, date):
                continue
            # 네비게이션 잔여 텍스트
            if line in ("받은 메모", "보낸 메모", "메모 쓰기", "받은 메모 보기",
                        "보낸 메모 보기", "메모 내용", "메모 보기"):
                continue
            body_lines.append(line)
        body = "\n".join(body_lines).strip()

    # 전부 비면 파싱 실패
    if not body.strip() and not sender and not recipient:
        return None

    return MemoContent(
        me_id="",
        sender=sender,
        recipient=recipient,
        date=date,
        body=body or "(본문 없음)",
        kind=kind,
    )


# ── API 호출 ──

def _is_login_redirect(resp: requests.Response) -> bool:
    """응답이 실제 로그인 페이지인지 정확히 판단.

    단순 "login.php" 문자열 포함은 오탐 — 인증된 페이지의 스크립트·네비게이션
    링크에도 login.php 가 들어갈 수 있다. 실제 로그인 페이지의 특징:
    1) 최종 URL 경로가 /bbs/login.php 로 끝남
    2) 페이지 title 이 "소리샘 로그인"
    3) form name="flogin" + mb_id + mb_password 입력 필드 존재
    """
    final_url = resp.url or ""
    # 경로가 login.php 로 끝나면 확실한 리다이렉트
    if re.search(r"/bbs/login\.php(\?|$)", final_url):
        return True
    # HTML 안에 로그인 폼이 실제로 있는지 확인
    head = resp.text[:4000]
    has_login_title = "<title>소리샘 로그인</title>" in head
    has_login_form = ('name="flogin"' in head or "flogin_submit" in head)
    has_password_input = 'name="mb_password"' in head
    return (has_login_title and has_password_input) or (has_login_form and has_password_input)


def _dump_debug_html(html: str, tag: str) -> str:
    """응답 HTML 을 data/memo_debug_<tag>.html 로 덤프. 디버깅 용."""
    import os
    from config import DATA_DIR
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        path = os.path.join(DATA_DIR, f"memo_debug_{tag}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        return path
    except OSError:
        return ""


def fetch_inbox(session: requests.Session, kind: str = "recv",
                page: int = 1) -> tuple[bool, list[MemoItem] | str]:
    """쪽지 목록 가져오기. kind='recv' 받은함, 'send' 보낸함. page=1 부터.
    성공 시 (True, items), 실패 시 (False, 오류메시지).
    """
    params = {"kind": kind}
    if page > 1:
        params["page"] = page
    try:
        resp = session.get(MEMO_LIST_URL, params=params, timeout=15)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}  (URL: {resp.url})"
    if _is_login_redirect(resp):
        return False, f"로그인 세션이 만료되었습니다. 다시 로그인해 주세요. (최종 URL: {resp.url})"
    # 목록 HTML 은 항상 덤프 — 개별 view 링크 구조 확인에 필요.
    # 성공/실패 관계없이 최근 1개만 유지.
    _dump_debug_html(resp.text, f"list_{kind}")
    items = _parse_memo_list(resp.text)
    if not items:
        dump_path = _dump_debug_html(resp.text, f"list_{kind}_empty")
        global _LAST_EMPTY_DUMP
        _LAST_EMPTY_DUMP = dump_path
    return True, items


_LAST_EMPTY_DUMP: str = ""


def fetch_inbox_up_to(session: requests.Session, kind: str = "recv",
                      target_count: int = 10, max_pages: int = 20
                      ) -> tuple[bool, list[MemoItem] | str]:
    """여러 페이지 순차 조회로 target_count 개까지 누적."""
    all_items: list = []
    seen_ids: set[str] = set()
    page = 1
    while len(all_items) < target_count and page <= max_pages:
        ok, items = fetch_inbox(session, kind=kind, page=page)
        if not ok:
            if page == 1:
                return False, items
            break
        if not items:
            break
        new_any = False
        for it in items:
            if it.me_id in seen_ids:
                continue
            seen_ids.add(it.me_id)
            all_items.append(it)
            new_any = True
            if len(all_items) >= target_count:
                break
        if not new_any:
            break
        page += 1
    return True, all_items[:target_count]


def get_last_empty_dump() -> str:
    """가장 최근 fetch_inbox 결과가 빈 목록이었을 때 덤프된 HTML 경로."""
    return _LAST_EMPTY_DUMP


def fetch_memo(session: requests.Session, me_id: str, kind: str = "recv") -> tuple[bool, MemoContent | str]:
    """개별 쪽지 읽기.

    ar.memo 플러그인은 `me_id` 만으로는 "값을 넘겨주세요" 에러를 돌려주므로
    `kind` 도 함께 보낸다. 혹시 그래도 실패할 경우 URL 파라미터 조합을
    순차 시도해 첫 성공 응답을 사용.
    """
    attempts = [
        {"me_id": me_id, "kind": kind},                    # 우선: me_id + kind
        {"me_id": me_id},                                   # 폴백 1: me_id 만 (gnuboard5 표준)
        {"me_id": me_id, "kind": kind, "mode": "view"},    # 폴백 2: mode=view 추가
        {"me_id": me_id, "mode": "view"},
    ]
    last_resp_text = ""
    last_url = ""
    last_err = ""

    for params in attempts:
        try:
            resp = session.get(MEMO_VIEW_URL, params=params, timeout=15)
        except requests.RequestException as e:
            last_err = f"서버 연결 실패: {e}"
            continue
        last_resp_text = resp.text
        last_url = resp.url
        if resp.status_code != 200:
            last_err = f"HTTP {resp.status_code}  (URL: {resp.url})"
            continue
        if _is_login_redirect(resp):
            last_err = f"로그인 세션이 만료되었습니다. (URL: {resp.url})"
            continue
        # 서버측 에러 페이지 감지 — alert("값을 넘겨주세요") 등
        if _looks_like_server_error_page(resp.text):
            last_err = "서버측 에러 응답 (파라미터 부족)"
            continue
        content = _parse_memo_content(resp.text, kind=kind)
        if content is not None:
            content.me_id = me_id
            # 성공 시에도 HTML 덤프 유지 (구조 변경 감지용, 최근 1개만)
            _dump_debug_html(resp.text, f"view_{me_id}")
            return True, content
        last_err = "파싱 실패"

    # 모든 시도 실패 — 마지막 응답 HTML 덤프
    dump_path = _dump_debug_html(last_resp_text, f"view_{me_id}")
    hint = f"  (디버그: {dump_path})" if dump_path else ""
    return False, f"쪽지 내용을 파싱하지 못했습니다. 최종 오류: {last_err}  (URL: {last_url}){hint}"


def _looks_like_server_error_page(html: str) -> bool:
    """gnuboard/ar.memo 의 'alert + history.back()' 에러 페이지 패턴 감지."""
    if "<title>오류안내 페이지</title>" in html:
        return True
    if "값을 넘겨주세요" in html:
        return True
    if "id=\"validation_check\"" in html and "history.back()" in html:
        return True
    return False


def send_memo(session: requests.Session, recipient: str, body: str) -> tuple[bool, str]:
    """쪽지 전송. gnuboard5 memo_form → memo_form_update 흐름.
    성공 시 (True, ""), 실패 시 (False, 오류메시지).
    """
    # 1. 폼 GET 해서 token 등 hidden 필드 수집
    try:
        form_resp = session.get(MEMO_FORM_URL, timeout=15)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if form_resp.status_code != 200 or _is_login_redirect(form_resp):
        return False, "쪽지 작성 폼을 불러올 수 없습니다. 로그인 상태를 확인해 주세요."

    soup = BeautifulSoup(form_resp.text, "lxml")
    form = soup.find("form", {"name": "fmemoform"}) or soup.find("form", id="fmemoform")
    if form is None:
        # 폴백 — 첫 form
        form = soup.find("form")
    if form is None:
        return False, "쪽지 작성 폼 구조를 인식하지 못했습니다."

    # 2. hidden 필드 수집
    data = {}
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        if inp.get("type") == "submit":
            continue
        data[name] = inp.get("value", "")

    # 3. 받는이·본문 설정
    data["me_recv_mb_id"] = recipient
    data["me_memo"] = body

    # 4. POST
    action = form.get("action") or MEMO_FORM_UPDATE_URL
    post_url = urljoin(form_resp.url, action)
    try:
        post_resp = session.post(post_url, data=data, timeout=15, allow_redirects=True)
    except requests.RequestException as e:
        return False, f"전송 실패: {e}"

    if post_resp.status_code != 200:
        return False, f"HTTP {post_resp.status_code}"

    # 5. gnuboard5 / ar.memo 는 alert() + location.href 로 응답.
    # 사이트마다 "전달되었습니다" / "전송되었습니다" / "발송되었습니다" 등 다름 →
    # 명시적 실패 키워드가 없으면 성공으로 간주.
    text = post_resp.text
    return _classify_alert_response(text, post_resp.url)


def _extract_immediate_alert_script(html: str) -> str:
    """함수 정의 내부가 아닌 즉시 실행되는 alert() 메시지 추출.

    gnuboard 목록 페이지는 JS 함수 안에 alert() 문자열이 코드로 존재하는데,
    이는 특정 조건에서만 실행되는 alert 라서 응답 해석에 쓰면 안 된다.
    function 키워드가 없는 <script> 블록의 alert 만 '즉시 실행' 으로 간주.
    """
    for script_match in re.finditer(
        r"<script[^>]*>(.*?)</script>", html, flags=re.DOTALL | re.I
    ):
        script = script_match.group(1)
        if re.search(r"function\s+\w+\s*\(", script) or re.search(
            r"=\s*function\s*\(", script
        ):
            continue
        m = re.search(r"alert\(\s*['\"]([^'\"]+)['\"]", script)
        if m:
            return m.group(1)
    return ""


def _classify_alert_response(text: str, final_url: str = "") -> tuple[bool, str]:
    """서버 응답 HTML 의 즉시 실행 alert() 해석해 (성공 여부, 메시지) 반환."""
    success_keywords = [
        "성공", "전송되었", "보냈", "발송되었", "전달되었",
        "전달하였", "보내졌", "완료"
    ]
    failure_keywords = [
        "실패", "오류", "잘못", "없습니다", "권한", "차단", "입력하",
        "거부", "이미 삭제", "존재하지 않"
    ]
    msg = _extract_immediate_alert_script(text)
    if msg:
        if any(w in msg for w in success_keywords):
            return True, msg
        if any(w in msg for w in failure_keywords):
            return False, msg
        return True, msg
    # alert 없음 — memo.php 로 리다이렉트됐으면 성공
    return True, ""


def delete_all_memos(session: requests.Session, kind: str = "recv") -> tuple[bool, str]:
    """쪽지함 전체 삭제 — ar.memo 의 memo_list_update.php 로 POST.

    서버측 폼 동작 (debug HTML 참조):
    - action: /plugin/ar.memo/memo_list_update.php
    - delete-function=delete-all
    - memo-type=recv|send
    - btn_submit=전체 삭제
    """
    data = {
        "delete-function": "delete-all",
        "memo-type": kind,
        "btn_submit": "전체 삭제",
    }
    try:
        resp = session.post(MEMO_LIST_UPDATE_URL, data=data, timeout=30,
                            allow_redirects=True)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}  (URL: {resp.url})"
    if _is_login_redirect(resp):
        return False, "로그인 세션이 만료되었습니다."
    return _classify_alert_response(resp.text, resp.url)


def _notify_log(msg: str):
    """쪽지 알림 디버그 로그 — data/memo_notify.log 에 append."""
    import os
    from datetime import datetime
    from config import DATA_DIR
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        path = os.path.join(DATA_DIR, "memo_notify.log")
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat()}] {msg}\n")
    except Exception:
        pass


def check_new_memos(session: requests.Session, timeout: float = 10.0) -> tuple[bool, str]:
    """/plugin/ar.memo/memo_check_new.php 호출 (레퍼런스용, 현재 Notifier 는 미사용).

    사이트 JS 기준: 응답 본문이 truthy 면 새 쪽지 있음.
    현재는 응답 포맷이 불확실해서 MemoNotifier 는 이 함수 대신 fetch_inbox 를
    직접 폴링해 seen_me_ids 와 비교한다 — 훨씬 견고함.
    """
    try:
        resp = session.get(MEMO_CHECK_NEW_URL, timeout=timeout)
    except requests.RequestException as e:
        _notify_log(f"check_new_memos exception: {e}")
        return False, ""
    if resp.status_code != 200:
        _notify_log(f"check_new_memos HTTP {resp.status_code}")
        return False, ""
    if _is_login_redirect(resp):
        _notify_log("check_new_memos got login redirect")
        return False, ""
    body = (resp.text or "").strip()
    _notify_log(f"check_new_memos body (len={len(body)}): {body[:200]!r}")
    if not body or body in ("0", "false", "null"):
        return False, body
    return True, body


class MemoNotifier:
    """백그라운드에서 주기적으로 새 쪽지 확인 후 콜백 호출.

    구조:
    - wx.Timer 가 UI 스레드에서 주기적 EVT_TIMER 발생
    - 각 tick 에서 별도 스레드로 fetch_inbox 수행 (UI 멈춤 방지)
    - seen_me_ids 에 없는 me_id 가 있으면 콜백 호출 (wx.CallAfter 로 UI 스레드)
    - 앱 시작 시 한 번 initial_fill — 기존 쪽지 전부 seen 처리 → 스팸 방지

    단순화 방침: memo_check_new.php 응답 포맷이 불확실해서 그 엔드포인트는
    사용하지 않고, 무조건 fetch_inbox 로 받은함 전체를 polling. 1분 간격 +
    응답 크기 작아 네트워크 비용 미미.
    """

    def __init__(self, parent_frame, session, callback_on_new):
        import wx as _wx
        self.frame = parent_frame
        self.session = session
        self.callback = callback_on_new
        self.seen_me_ids: set[str] = set()
        self._in_flight = False
        self._initial_done = False
        self.timer = _wx.Timer(parent_frame)
        parent_frame.Bind(_wx.EVT_TIMER, self._on_tick, self.timer)
        _notify_log("MemoNotifier created")

    def start(self, interval_sec: int = MEMO_NOTIFY_INTERVAL_SEC):
        """초기 seen 채우기(백그라운드) + 타이머 시작."""
        _notify_log(f"start(interval={interval_sec}s) — initial_fill in bg")
        import threading
        def init_worker():
            try:
                ok, items = fetch_inbox(self.session, kind="recv")
                if ok:
                    for it in items:
                        self.seen_me_ids.add(it.me_id)
                    _notify_log(f"initial_fill OK: {len(items)} items marked seen")
                else:
                    _notify_log(f"initial_fill FAIL: {items}")
            except Exception as e:
                _notify_log(f"initial_fill exception: {e}")
            finally:
                self._initial_done = True
        threading.Thread(target=init_worker, daemon=True).start()
        self.timer.Start(interval_sec * 1000)
        _notify_log(f"timer started (IsRunning={self.timer.IsRunning()})")

    def stop(self):
        self.timer.Stop()
        _notify_log("stopped")

    # 호환성: 기존 호출자는 initial_fill_async 를 호출함 — 이제 start 에 통합.
    def initial_fill_async(self):
        pass

    def _on_tick(self, event):
        _notify_log(f"_on_tick fired (in_flight={self._in_flight}, initial_done={self._initial_done})")
        if not self._initial_done:
            return
        if self._in_flight:
            return
        import threading
        self._in_flight = True
        threading.Thread(target=self._check_in_bg, daemon=True).start()

    def _check_in_bg(self):
        import wx as _wx
        try:
            ok, items = fetch_inbox(self.session, kind="recv")
            if not ok:
                _notify_log(f"fetch_inbox failed: {items}")
                return
            new_items = [it for it in items if it.me_id not in self.seen_me_ids]
            _notify_log(f"check: total={len(items)} seen={len(self.seen_me_ids)} new={len(new_items)}")
            if not new_items:
                return
            for it in new_items:
                self.seen_me_ids.add(it.me_id)
            _notify_log(f"→ alerting {len(new_items)} new item(s)")
            _wx.CallAfter(self.callback, len(new_items), new_items)
        except Exception as e:
            _notify_log(f"_check_in_bg exception: {e}")
        finally:
            self._in_flight = False

    def check_now(self, on_no_new=None, on_error=None):
        """수동 트리거 — 즉시 체크. 자동 틱과 달리 결과를 피드백.

        on_no_new: 새 쪽지가 없을 때 UI 스레드에서 호출될 콜백 (선택)
        on_error:  오류 발생 시 UI 스레드에서 호출될 콜백 (선택), 인자: str 메시지
        """
        _notify_log("check_now() manual trigger")
        if not self._initial_done:
            _notify_log("check_now: initial fill not done yet, retrying in 2s")
            import wx as _wx
            _wx.CallLater(2000, lambda: self.check_now(on_no_new, on_error))
            return
        import threading
        threading.Thread(
            target=self._manual_check_in_bg,
            args=(on_no_new, on_error),
            daemon=True,
        ).start()

    def _manual_check_in_bg(self, on_no_new, on_error):
        """수동 체크 — seen_me_ids 필터링 무시하고 현재 '안 읽은' 쪽지 전체를 알림.

        자동 폴링과 달리 사용자가 직접 요청한 확인이므로:
        - is_read==False 인 쪽지를 모두 unread 로 간주
        - 이미 seen 에 있더라도 아직 안 읽은 상태면 알림 대상
        - 자동 폴링의 중복 스팸 방지는 유지 (seen 집합은 건드리지 않음)
        """
        import wx as _wx
        try:
            ok, items = fetch_inbox(self.session, kind="recv")
            if not ok:
                _notify_log(f"manual fetch_inbox failed: {items}")
                if on_error:
                    _wx.CallAfter(on_error, str(items))
                return
            unread_items = [it for it in items if not it.is_read]
            _notify_log(
                f"manual check: total={len(items)} unread={len(unread_items)} "
                f"(seen_set_size={len(self.seen_me_ids)})"
            )
            if not unread_items:
                if on_no_new:
                    _wx.CallAfter(on_no_new)
                return
            # 수동 체크는 seen 집합을 수정하지 않음 — 사용자가 다시 Ctrl+Shift+N
            # 누르면 또 동일한 unread 항목들을 보여줘야 하기 때문.
            _notify_log(f"→ manual alerting {len(unread_items)} unread item(s)")
            _wx.CallAfter(self.callback, len(unread_items), unread_items)
        except Exception as e:
            _notify_log(f"_manual_check_in_bg exception: {e}")
            if on_error:
                _wx.CallAfter(on_error, str(e))

    def mark_all_as_seen(self):
        """사용자가 받은 쪽지함을 직접 열어서 확인한 뒤 호출 —
        현재 받은함의 모든 me_id 를 seen 으로 업데이트."""
        import threading
        def worker():
            try:
                ok, items = fetch_inbox(self.session, kind="recv")
                if ok:
                    for it in items:
                        self.seen_me_ids.add(it.me_id)
                    _notify_log(f"mark_all_as_seen: {len(items)} items")
            except Exception as e:
                _notify_log(f"mark_all_as_seen exception: {e}")
        threading.Thread(target=worker, daemon=True).start()


def delete_memo(session: requests.Session, me_id: str, kind: str = "recv") -> tuple[bool, str]:
    """쪽지 삭제. gnuboard5 memo_delete.php 는 GET. token 이 필요할 수 있음."""
    params = {"me_id": me_id, "kind": kind}
    try:
        # memo_view 에서 token 을 가져오는 흐름이 있을 수 있으니 먼저 view 접근
        view_resp = session.get(MEMO_VIEW_URL, params={"me_id": me_id}, timeout=15)
        if view_resp.status_code == 200 and not _is_login_redirect(view_resp):
            soup = BeautifulSoup(view_resp.text, "lxml")
            token_input = soup.find("input", {"name": "token"})
            if token_input and token_input.get("value"):
                params["token"] = token_input.get("value")
            # 또는 "삭제" 링크의 href 에서 직접 token 파라미터 찾기
            del_a = soup.find("a", href=re.compile(r"memo_delete\.php"))
            if del_a:
                m = re.search(r"token=([A-Za-z0-9]+)", del_a.get("href", ""))
                if m:
                    params["token"] = m.group(1)
    except requests.RequestException:
        pass

    try:
        resp = session.get(MEMO_DELETE_URL, params=params, timeout=15, allow_redirects=True)
    except requests.RequestException as e:
        return False, f"삭제 요청 실패: {e}"

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    return _classify_alert_response(resp.text, resp.url)


# ── UI 대화상자 ──

class _MemoReadTextCtrl(wx.TextCtrl):
    """스크린리더 친화 read-only TextCtrl — 한 줄 항목용."""

    def MSWHandleMessage(self, msg, wParam, lParam):
        WM_KEYDOWN = 0x0100
        WM_CHAR = 0x0102
        if msg in (WM_KEYDOWN, WM_CHAR):
            # Home/End/Up/Down/Return/Back/PageUp/PageDn/Escape/Delete
            blocked_vk = {0x24, 0x23, 0x26, 0x28, 0x0D, 0x08, 0x21, 0x22, 0x1B, 0x2E}
            if wParam in blocked_vk:
                return True, 0
        return super().MSWHandleMessage(msg, wParam, lParam)


class MemoViewDialog(wx.Dialog):
    """개별 쪽지 보기.

    items + index 를 받으면 PageUp/PageDown 과 이전/다음 버튼으로 쪽지 사이를
    바로 이동 가능 (목록으로 돌아갈 필요 없음).
    """

    def __init__(self, parent, session: requests.Session, content: MemoContent,
                 items: list | None = None, index: int = 0,
                 kind: str | None = None):
        kind_label = "받은 쪽지" if content.kind == "recv" else "보낸 쪽지"
        super().__init__(parent, title=kind_label, size=(720, 540),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.content = content
        self.items = items or []
        self.index = index
        self.kind = kind or content.kind

        panel = wx.Panel(self)
        self._panel = panel
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 메타+본문 통합 표시용 단일 read-only TextCtrl.
        # 스크린리더가 아래 방향키로 작성 날짜 → 작성 시간 → 보낸/받는 사람 →
        # 빈 줄 → "내용:" → 본문 순으로 자연스럽게 읽어 내려갈 수 있게 한 영역에 통합.
        self.body_ctrl = wx.TextCtrl(
            panel, value="",
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
        )
        vbox.Add(self.body_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.prev_btn = None
        self.next_btn = None
        if self.items:
            self.prev_btn = wx.Button(panel, label="이전(&P)")
            self.prev_btn.Bind(wx.EVT_BUTTON, lambda e: self._navigate(-1))
            btn_sizer.Add(self.prev_btn, 0, wx.RIGHT, 8)
            self.next_btn = wx.Button(panel, label="다음(&N)")
            self.next_btn.Bind(wx.EVT_BUTTON, lambda e: self._navigate(+1))
            btn_sizer.Add(self.next_btn, 0, wx.RIGHT, 8)
        if self.kind == "recv":
            self.reply_btn = wx.Button(panel, label="답장(&R)")
            self.reply_btn.Bind(wx.EVT_BUTTON, self.on_reply)
            btn_sizer.Add(self.reply_btn, 0, wx.RIGHT, 8)
        self.delete_btn = wx.Button(panel, label="삭제(&D)")
        self.delete_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_sizer.Add(self.delete_btn, 0, wx.RIGHT, 8)
        self.close_btn = wx.Button(panel, wx.ID_CLOSE, label="닫기(&C)")
        self.close_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(wx.ID_CLOSE))
        btn_sizer.Add(self.close_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)

        # 테마 적용
        apply_theme(self, make_font(load_font_size()))

        # 단축키
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        # 초기 표시
        self._refresh_display(announce=False)

        # 본문으로 포커스
        self.body_ctrl.SetFocus()
        self.body_ctrl.SetInsertionPoint(0)

    def _refresh_display(self, announce: bool = True):
        """현재 self.content 기준으로 meta·body·버튼 상태 갱신."""
        kind_label = "받은 쪽지" if self.content.kind == "recv" else "보낸 쪽지"
        total = len(self.items)
        if total:
            self.SetTitle(f"{kind_label} ({self.index + 1}/{total})")
        else:
            self.SetTitle(kind_label)

        # 날짜와 시간을 분리 ("26-04-23 13:47" → 날짜 "26-04-23" / 시간 "13:47")
        date_part = ""
        time_part = ""
        if self.content.date:
            parts = self.content.date.strip().split(None, 1)
            if len(parts) >= 2:
                date_part, time_part = parts[0], parts[1]
            else:
                if ":" in self.content.date:
                    time_part = self.content.date
                else:
                    date_part = self.content.date

        # 사람 라벨 자동 전환
        if self.content.kind == "recv":
            person_label = "보낸 사람"
            person_value = self.content.sender or "(알 수 없음)"
        else:
            person_label = "받는 사람"
            person_value = self.content.recipient or "(알 수 없음)"

        # 본문 영역에 "작성 날짜 → 작성 시간 → 사람 → 빈 줄 → 내용: → 본문" 통합
        lines = [
            f"작성 날짜: {date_part or '(정보 없음)'}",
            f"작성 시간: {time_part or '(정보 없음)'}",
            f"{person_label}: {person_value}",
            "",
            "내용:",
            self.content.body or "(본문 없음)",
        ]
        self.body_ctrl.SetValue("\n".join(lines))
        self.body_ctrl.SetInsertionPoint(0)

        # 이전/다음 버튼 활성 상태 (버튼 없어도 None 체크)
        if self.prev_btn:
            self.prev_btn.Enable(self.index > 0)
        if self.next_btn:
            self.next_btn.Enable(self.index < len(self.items) - 1)

        if hasattr(self, "_panel"):
            self._panel.Layout()

        if announce and total:
            speak(f"{self.index + 1} / {total} · {person_value}")

    def _navigate(self, direction: int):
        """direction: -1 = 이전, +1 = 다음. 경계에 가면 음성 안내만."""
        if not self.items:
            return
        new_idx = self.index + direction
        if new_idx < 0:
            speak("첫 번째 쪽지입니다.")
            return
        if new_idx >= len(self.items):
            speak("마지막 쪽지입니다.")
            return
        item = self.items[new_idx]
        speak("쪽지를 불러옵니다.")
        ok, result = fetch_memo(self.session, item.me_id, kind=self.kind)
        if not ok:
            speak("쪽지를 불러오지 못했습니다.")
            wx.MessageBox(f"쪽지를 불러오지 못했습니다.\n{result}",
                          "오류", wx.OK | wx.ICON_ERROR, self)
            return
        self.index = new_idx
        self.content = result
        # 방향 안내 + 위치·상대방 정보 음성
        dir_msg = "다음 쪽지입니다." if direction > 0 else "이전 쪽지입니다."
        self._refresh_display(announce=False)
        who = self.content.sender or self.content.recipient or "알 수 없음"
        speak(f"{dir_msg} {self.index + 1} / {len(self.items)} · {who}")
        # 본문으로 포커스 재설정
        self.body_ctrl.SetFocus()
        self.body_ctrl.SetInsertionPoint(0)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        alt_only = event.AltDown() and not event.ControlDown() and not event.ShiftDown()

        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CLOSE)
            return
        if (key == ord("D") or key == wx.WXK_DELETE) and not mods:
            self.on_delete(None)
            return
        if key == ord("R") and not mods and self.content.kind == "recv":
            self.on_reply(None)
            return
        # PageUp/PageDown 또는 Alt+P / Alt+N 로 이전/다음 쪽지 이동
        if key == wx.WXK_PAGEUP and not mods:
            self._navigate(-1)
            return
        if key == wx.WXK_PAGEDOWN and not mods:
            self._navigate(+1)
            return
        if key == ord("P") and alt_only:
            self._navigate(-1)
            return
        if key == ord("N") and alt_only:
            self._navigate(+1)
            return
        event.Skip()

    def on_reply(self, event):
        dlg = MemoWriteDialog(self, self.session,
                              default_recipient=self.content.sender,
                              default_body=f"\n\n--- 원본 쪽지 ---\n{self.content.body}")
        dlg.ShowModal()
        dlg.Destroy()

    def on_delete(self, event):
        ans = wx.MessageBox("이 쪽지를 삭제하시겠습니까?",
                            "쪽지 삭제", wx.YES_NO | wx.ICON_QUESTION, self)
        if ans != wx.YES:
            return
        ok, msg = delete_memo(self.session, self.content.me_id, kind=self.content.kind)
        if ok:
            speak("쪽지를 삭제했습니다.")
            # 삭제 후: items 리스트에서 현재 항목 제거하고 다음(또는 이전) 으로 이동
            if self.items:
                del self.items[self.index]
                if not self.items:
                    wx.MessageBox("마지막 쪽지였습니다. 쪽지함으로 돌아갑니다.",
                                  "삭제 완료", wx.OK | wx.ICON_INFORMATION, self)
                    self.EndModal(wx.ID_OK)
                    return
                # 가능하면 같은 index 유지, 마지막이었다면 한 칸 앞으로
                if self.index >= len(self.items):
                    self.index = len(self.items) - 1
                # 다음 쪽지 로드
                item = self.items[self.index]
                ok2, result = fetch_memo(self.session, item.me_id, kind=self.kind)
                if ok2:
                    self.content = result
                    self._refresh_display(announce=True)
                    self.body_ctrl.SetFocus()
                    self.body_ctrl.SetInsertionPoint(0)
                    return
                # 불러오기 실패 — 목록으로 복귀
                self.EndModal(wx.ID_OK)
                return
            # items 없는 경우 (구 호출자) — 기존 동작 유지
            wx.MessageBox("쪽지를 삭제했습니다.", "삭제 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        else:
            speak("쪽지 삭제 실패.")
            wx.MessageBox(f"쪽지 삭제에 실패했습니다.\n{msg}",
                          "삭제 실패", wx.OK | wx.ICON_ERROR, self)


class MemoWriteDialog(wx.Dialog):
    """쪽지 작성 / 답장."""

    def __init__(self, parent, session: requests.Session,
                 default_recipient: str = "", default_body: str = ""):
        super().__init__(parent, title="쪽지 작성", size=(640, 480),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 받는이
        lbl_r = wx.StaticText(panel, label="받는 사람 아이디")
        vbox.Add(lbl_r, 0, wx.TOP | wx.LEFT | wx.RIGHT, 8)
        self.recipient_ctrl = wx.TextCtrl(panel, value=default_recipient)
        vbox.Add(self.recipient_ctrl, 0, wx.ALL | wx.EXPAND, 8)

        # 본문
        lbl_b = wx.StaticText(panel, label="쪽지 내용")
        vbox.Add(lbl_b, 0, wx.LEFT | wx.RIGHT, 8)
        self.body_ctrl = wx.TextCtrl(panel, value=default_body,
                                     style=wx.TE_MULTILINE | wx.TE_RICH2)
        vbox.Add(self.body_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.send_btn = wx.Button(panel, label="보내기(&S)")
        self.send_btn.Bind(wx.EVT_BUTTON, self.on_send)
        btn_sizer.Add(self.send_btn, 0, wx.RIGHT, 8)
        self.cancel_btn = wx.Button(panel, wx.ID_CANCEL, label="취소")
        btn_sizer.Add(self.cancel_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))

        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        if default_recipient:
            self.body_ctrl.SetFocus()
        else:
            self.recipient_ctrl.SetFocus()

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key == ord("S") and event.ControlDown():
            self.on_send(None)
            return
        event.Skip()

    def on_send(self, event):
        recipient = self.recipient_ctrl.GetValue().strip()
        body = self.body_ctrl.GetValue().strip()
        if not recipient:
            wx.MessageBox("받는 사람 아이디를 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.recipient_ctrl.SetFocus()
            return
        if not body:
            wx.MessageBox("쪽지 내용을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.body_ctrl.SetFocus()
            return

        speak("쪽지를 전송하는 중입니다.")
        self.send_btn.Disable()
        ok, msg = send_memo(self.session, recipient, body)
        self.send_btn.Enable()

        if ok:
            speak("쪽지를 보냈습니다.")
            wx.MessageBox(msg or "쪽지를 보냈습니다.", "전송 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        else:
            speak("쪽지 전송에 실패했습니다.")
            wx.MessageBox(f"쪽지 전송에 실패했습니다.\n{msg}", "전송 실패",
                          wx.OK | wx.ICON_ERROR, self)


class MemoInboxDialog(wx.Dialog):
    """쪽지함 — 받은/보낸 전환 + 목록 + 개별 보기 연결."""

    def __init__(self, parent, session: requests.Session):
        super().__init__(parent, title="쪽지함", size=(760, 560),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.kind = "recv"
        self.items: list[MemoItem] = []
        self._index = 0

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 받은/보낸 전환 버튼
        top = wx.BoxSizer(wx.HORIZONTAL)
        self.recv_btn = wx.Button(panel, label="받은 쪽지함(&R)")
        self.recv_btn.Bind(wx.EVT_BUTTON, lambda e: self._switch("recv"))
        self.send_btn = wx.Button(panel, label="보낸 쪽지함(&S)")
        self.send_btn.Bind(wx.EVT_BUTTON, lambda e: self._switch("send"))
        self.compose_btn = wx.Button(panel, label="새 쪽지(&N)")
        self.compose_btn.Bind(wx.EVT_BUTTON, self.on_compose)
        self.refresh_btn = wx.Button(panel, label="새로고침(&F)")
        self.refresh_btn.Bind(wx.EVT_BUTTON, lambda e: self.reload())
        self.delete_all_btn = wx.Button(panel, label="모든 쪽지 삭제(&A)")
        self.delete_all_btn.Bind(wx.EVT_BUTTON, self.on_delete_all)
        top.Add(self.recv_btn, 0, wx.RIGHT, 4)
        top.Add(self.send_btn, 0, wx.RIGHT, 4)
        top.Add(self.compose_btn, 0, wx.RIGHT, 4)
        top.Add(self.refresh_btn, 0, wx.RIGHT, 4)
        top.Add(self.delete_all_btn, 0)
        vbox.Add(top, 0, wx.ALL, 8)

        # 상태 라벨
        self.status_label = wx.StaticText(panel, label="불러오는 중...")
        vbox.Add(self.status_label, 0, wx.LEFT | wx.RIGHT, 8)

        # 쪽지 목록 (ListBox — 스크린리더가 선택 항목을 자동 낭독)
        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._open_current())
        vbox.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 키 안내
        hint = wx.StaticText(
            panel,
            label="↑↓ 이동 · PgUp/PgDn 더 불러오기 · Enter 읽기 · D/Del 삭제 · Shift+Del 전체 삭제 · R 답장 · N 새 쪽지 · Esc 닫기"
        )
        vbox.Add(hint, 0, wx.ALL, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))

        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        self.list_ctrl.SetFocus()

        wx.CallAfter(self.reload)

    def _switch(self, kind: str):
        if kind == self.kind:
            return
        self.kind = kind
        self.reload()

    def reload(self):
        self.status_label.SetLabel(
            f"{'받은' if self.kind == 'recv' else '보낸'} 쪽지함 불러오는 중..."
        )
        # 사용자 설정의 표시 개수까지 페이징해서 가져옴
        try:
            from settings_dialog import load_notify_settings
            target = int(load_notify_settings().get("list_page_size", 10))
        except Exception:
            target = 10
        self._target_per_reload = target
        # 몇 페이지까지 로드했는지 추적 (1페이지 당 약 target 개)
        self._loaded_pages = 1
        ok, result = fetch_inbox_up_to(self.session, kind=self.kind, target_count=target)
        if not ok:
            self.items = []
            self._index = 0
            self.list_ctrl.Clear()
            self.status_label.SetLabel(f"오류: {result}")
            speak(f"쪽지함을 불러오지 못했습니다.")
            wx.MessageBox(
                f"쪽지함을 불러오지 못했습니다.\n\n"
                f"원인: {result}\n\n"
                f"세션이 만료되었다면 프로그램을 다시 시작해 주세요.",
                "쪽지함 불러오기 실패",
                wx.OK | wx.ICON_ERROR, self,
            )
            return
        self.items = result
        self._index = 0
        label = "받은" if self.kind == "recv" else "보낸"
        if not self.items:
            self.list_ctrl.Clear()
            dump_path = get_last_empty_dump()
            if dump_path:
                self.status_label.SetLabel(
                    f"{label} 쪽지가 없거나 파싱 실패 — 디버그: {dump_path}"
                )
            else:
                self.status_label.SetLabel(f"{label} 쪽지가 없습니다.")
            speak(f"{label} 쪽지가 없습니다.")
            return
        self.status_label.SetLabel(f"{label} 쪽지 {len(self.items)}개")
        self._update_display()
        speak(f"{label} 쪽지 {len(self.items)}개.")

    def _format_item(self, i: int, item: MemoItem) -> str:
        label_who = "보낸이" if self.kind == "recv" else "받는이"
        unread = " [안읽음]" if self.kind == "recv" and not item.is_read else ""
        return f"{i+1}/{len(self.items)}{unread} · {label_who}: {item.counterpart} · {item.date} · {item.summary}"

    def _update_display(self):
        if not self.items:
            self.list_ctrl.Clear()
            return
        self._index = max(0, min(self._index, len(self.items) - 1))
        lines = [self._format_item(i, it) for i, it in enumerate(self.items)]
        self.list_ctrl.Set(lines)
        self.list_ctrl.SetSelection(self._index)
        self.list_ctrl.SetFocus()

    def _sync_index_from_listbox(self):
        """ListBox 의 현재 선택을 self._index 에 반영."""
        sel = self.list_ctrl.GetSelection()
        if sel != wx.NOT_FOUND:
            self._index = sel

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()

        # 포커스가 버튼에 있으면 버튼 기본 동작 우선
        focused = self.FindFocus()
        if isinstance(focused, wx.Button):
            event.Skip()
            return

        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CLOSE)
            return
        if key == wx.WXK_RETURN and not mods:
            self._sync_index_from_listbox()
            self._open_current()
            return
        # Shift+Del 은 전체 삭제 (D/Del 단독보다 먼저 검사)
        if key == wx.WXK_DELETE and event.ShiftDown() and not event.ControlDown() and not event.AltDown():
            self.on_delete_all(None)
            return
        if (key == ord("D") or key == wx.WXK_DELETE) and not mods:
            self._sync_index_from_listbox()
            self._delete_current()
            return
        if key == ord("R") and not mods and self.kind == "recv":
            self._sync_index_from_listbox()
            self._reply_current()
            return
        if key == ord("N") and not mods:
            self.on_compose(None)
            return
        if key == ord("F") and not mods:
            self.reload()
            return
        # PageUp/PageDown — 다음 페이지를 누적 로드 (현재 리스트에 append)
        if key in (wx.WXK_PAGEDOWN, wx.WXK_PAGEUP) and not mods:
            self._load_more()
            return
        # ↑↓ 등 기타 키는 ListBox 네이티브 동작에 맡김 (스크린리더가 자동 낭독)
        event.Skip()

    def _load_more(self):
        """다음 페이지를 추가로 불러와 현재 목록에 append."""
        if not hasattr(self, "_loaded_pages"):
            self._loaded_pages = 1
        if not hasattr(self, "_target_per_reload"):
            try:
                from settings_dialog import load_notify_settings
                self._target_per_reload = int(load_notify_settings().get("list_page_size", 10))
            except Exception:
                self._target_per_reload = 10

        next_page = self._loaded_pages + 1
        speak("다음 쪽지를 불러옵니다.")
        ok, new_items = fetch_inbox(self.session, kind=self.kind, page=next_page)
        if not ok:
            speak("불러오기에 실패했습니다.")
            return
        if not new_items:
            speak("더 이상 불러올 쪽지가 없습니다.")
            return

        # 중복 제거하며 누적
        existing_ids = {it.me_id for it in self.items}
        added = []
        for it in new_items:
            if it.me_id in existing_ids:
                continue
            existing_ids.add(it.me_id)
            self.items.append(it)
            added.append(it)

        if not added:
            speak("더 이상 불러올 쪽지가 없습니다.")
            return

        self._loaded_pages = next_page

        # ListBox 갱신 — 새 항목만 Append (기존 선택 유지)
        sel = self.list_ctrl.GetSelection()
        for i, it in enumerate(added, start=len(self.items) - len(added)):
            self.list_ctrl.Append(self._format_item(i, it))
        # 기존 항목 번호도 "N/total" 이 바뀌었으니 전체 재포맷
        lines = [self._format_item(i, it) for i, it in enumerate(self.items)]
        self.list_ctrl.Set(lines)
        if sel == wx.NOT_FOUND:
            # 새로 추가된 첫 항목으로 포커스
            self.list_ctrl.SetSelection(len(self.items) - len(added))
        else:
            self.list_ctrl.SetSelection(sel)

        label = "받은" if self.kind == "recv" else "보낸"
        self.status_label.SetLabel(f"{label} 쪽지 {len(self.items)}개 (페이지 {self._loaded_pages}까지)")
        speak(f"{len(added)}개 추가로 불러왔습니다. 총 {len(self.items)}개.")

    def _open_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        speak("쪽지를 불러옵니다.")
        ok, result = fetch_memo(self.session, item.me_id, kind=self.kind)
        if not ok:
            speak("쪽지를 불러오지 못했습니다.")
            wx.MessageBox(f"쪽지를 불러오지 못했습니다.\n{result}",
                          "오류", wx.OK | wx.ICON_ERROR, self)
            return
        # items + index + kind 를 전달해서 PageUp/PageDown 네비게이션 활성화
        dlg = MemoViewDialog(self, self.session, result,
                             items=self.items, index=self._index, kind=self.kind)
        code = dlg.ShowModal()
        final_index = dlg.index
        dlg.Destroy()
        if code == wx.ID_OK:
            # 삭제됨 — 목록 갱신
            self.reload()
        else:
            # 네비게이션으로 마지막에 선택됐던 위치를 목록에 반영
            if self.items and 0 <= final_index < len(self.items):
                self._index = final_index
                self.list_ctrl.SetSelection(self._index)

    def _delete_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        ans = wx.MessageBox(f"선택한 쪽지를 삭제하시겠습니까?\n\n{item.summary}",
                            "쪽지 삭제", wx.YES_NO | wx.ICON_QUESTION, self)
        if ans != wx.YES:
            return
        ok, msg = delete_memo(self.session, item.me_id, kind=self.kind)
        if ok:
            speak("쪽지를 삭제했습니다.")
            self.reload()
        else:
            speak("쪽지 삭제에 실패했습니다.")
            wx.MessageBox(f"삭제에 실패했습니다.\n{msg}", "오류",
                          wx.OK | wx.ICON_ERROR, self)

    def _reply_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        dlg = MemoWriteDialog(self, self.session,
                              default_recipient=item.counterpart)
        dlg.ShowModal()
        dlg.Destroy()

    def on_compose(self, event):
        dlg = MemoWriteDialog(self, self.session)
        code = dlg.ShowModal()
        dlg.Destroy()
        if code == wx.ID_OK and self.kind == "send":
            self.reload()

    def on_delete_all(self, event):
        """현재 쪽지함(받은/보낸)의 모든 쪽지를 일괄 삭제."""
        label = "받은" if self.kind == "recv" else "보낸"
        if not self.items:
            speak(f"{label} 쪽지함이 이미 비어 있습니다.")
            wx.MessageBox(f"{label} 쪽지함에 삭제할 쪽지가 없습니다.",
                          "알림", wx.OK | wx.ICON_INFORMATION, self)
            return
        ans = wx.MessageBox(
            f"{label} 쪽지함의 전체 쪽지 {len(self.items)}개를 모두 삭제하시겠습니까?\n"
            "삭제한 쪽지는 복구할 수 없습니다.",
            "전체 쪽지 삭제",
            wx.YES_NO | wx.ICON_WARNING | wx.NO_DEFAULT, self,
        )
        if ans != wx.YES:
            return
        speak("전체 쪽지를 삭제하는 중입니다.")
        ok, msg = delete_all_memos(self.session, kind=self.kind)
        if ok:
            speak(f"{label} 쪽지함을 비웠습니다.")
            wx.MessageBox(msg or f"{label} 쪽지함의 모든 쪽지를 삭제했습니다.",
                          "삭제 완료", wx.OK | wx.ICON_INFORMATION, self)
            self.reload()
        else:
            speak("전체 쪽지 삭제 실패.")
            wx.MessageBox(f"전체 쪽지 삭제에 실패했습니다.\n{msg}",
                          "삭제 실패", wx.OK | wx.ICON_ERROR, self)
