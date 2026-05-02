"""DAISY 도서 압축 해제 + DTBook XML → TXT 변환 (v1.7).

DAISY 도서는 보통 ZIP 으로 묶여 배포되며, 안에 .opf / .ncx / .xml(=DTBook) /
.smil / .mp3 등이 들어 있다. 이 모듈은:

  · ZIP 파일이 DAISY 인지 자동 판정
  · ZIP 을 같은 이름의 폴더로 풀어내고
  · DTBook XML 의 본문을 평문 TXT 로 변환 (NCX 의 목차도 머리에 첨부)
"""

from __future__ import annotations

import os
import re
import zipfile

from bs4 import BeautifulSoup


# ── 판정 ──

def is_daisy_zip(path: str) -> bool:
    """ZIP 파일 안에 DAISY 시그너처(.opf / .ncx) 가 있으면 True."""
    if not path or not path.lower().endswith(".zip"):
        return False
    if not os.path.isfile(path):
        return False
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                low = name.lower()
                if low.endswith(".opf") or low.endswith(".ncx"):
                    return True
                if low.endswith(".xml") and ("dtbook" in low or "book" in low):
                    return True
    except (zipfile.BadZipFile, OSError):
        return False
    return False


# ── 압축 해제 ──

def extract_zip(path: str, dest_dir: str | None = None) -> str:
    """ZIP 을 dest_dir 에 풀고 추출 폴더 경로 반환.

    dest_dir 이 None 이면 ZIP 과 같은 위치에 'ZIP 이름' 폴더로 풀어낸다.
    이미 같은 이름 폴더가 있으면 그대로 사용 (덮어쓰기).
    """
    base = os.path.basename(path)
    name_no_ext = os.path.splitext(base)[0]
    if dest_dir is None:
        parent = os.path.dirname(os.path.abspath(path))
        dest_dir = os.path.join(parent, name_no_ext)
    os.makedirs(dest_dir, exist_ok=True)
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(dest_dir)
    return dest_dir


# ── DTBook XML → 평문 ──

_BLOCK_TAGS = {
    "p", "div", "li", "tr", "br", "h1", "h2", "h3", "h4", "h5", "h6",
    "level1", "level2", "level3", "level4", "level5", "level6",
    "doctitle", "docauthor", "frontmatter", "bodymatter", "rearmatter",
    "imggroup", "blockquote", "note", "noteref", "sidebar",
}


def _walk_text(node, out: list[str]) -> None:
    """DTBook XML 트리를 깊이 우선으로 돌면서 텍스트 수집.

    블록 태그 사이에는 줄바꿈을 넣어 단락 구분이 살아 있도록 한다.
    """
    from bs4 import NavigableString, Tag
    if isinstance(node, NavigableString):
        s = str(node)
        if s:
            out.append(s)
        return
    if not isinstance(node, Tag):
        return
    name = node.name.lower() if node.name else ""
    is_block = name in _BLOCK_TAGS
    if is_block and out and not out[-1].endswith("\n"):
        out.append("\n")
    if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
        # 헤딩은 빈 줄과 #으로 강조
        out.append("\n")
        level = int(name[1])
        out.append("#" * level + " ")
    for child in node.children:
        _walk_text(child, out)
    if is_block and out and not out[-1].endswith("\n"):
        out.append("\n")


def dtbook_to_text(xml_path: str) -> str:
    """DTBook XML 파일을 평문 텍스트로 변환."""
    with open(xml_path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    soup = BeautifulSoup(raw, "lxml-xml")
    body = soup.find("bodymatter") or soup.find("body") or soup.find("dtbook") or soup
    out: list[str] = []
    _walk_text(body, out)
    text = "".join(out)
    # 공백 정리
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


# ── NCX → 목차 ──

def ncx_to_toc(ncx_path: str) -> str:
    """NCX 파일에서 목차를 추출해 들여쓰기된 텍스트로 반환."""
    try:
        with open(ncx_path, "r", encoding="utf-8", errors="replace") as f:
            raw = f.read()
    except OSError:
        return ""
    soup = BeautifulSoup(raw, "lxml-xml")
    out: list[str] = []

    def walk(nav, level: int):
        for np in nav.find_all("navPoint", recursive=False):
            label = np.find("navLabel")
            text = label.get_text(strip=True) if label else ""
            if text:
                out.append(("  " * level) + "- " + text)
            walk(np, level + 1)

    nav_map = soup.find("navMap")
    if nav_map:
        walk(nav_map, 0)
    return "\n".join(out)


# ── 메타 ──

def book_title(opf_or_ncx_path: str) -> str:
    """OPF 또는 NCX 에서 도서 제목을 추출. 실패하면 빈 문자열."""
    try:
        with open(opf_or_ncx_path, "r", encoding="utf-8", errors="replace") as f:
            raw = f.read()
    except OSError:
        return ""
    soup = BeautifulSoup(raw, "lxml-xml")
    el = (
        soup.find("dc:title")
        or soup.find("docTitle")
        or soup.find("title")
    )
    if el:
        return el.get_text(strip=True)
    return ""


# ── 통합 ──

def find_dtbook_xml(folder: str) -> str | None:
    """폴더 안에서 DTBook 본문 XML 파일을 찾아 경로 반환."""
    candidates = []
    for root, _, files in os.walk(folder):
        for fn in files:
            low = fn.lower()
            if not low.endswith(".xml"):
                continue
            full = os.path.join(root, fn)
            # DTBook root 태그가 있는지 가볍게 확인
            try:
                with open(full, "r", encoding="utf-8", errors="replace") as f:
                    head = f.read(2048)
            except OSError:
                continue
            if "<dtbook" in head.lower() or "<book" in head.lower():
                candidates.append(full)
    if candidates:
        # 가장 큰 파일을 본문으로 추정
        candidates.sort(key=lambda p: os.path.getsize(p), reverse=True)
        return candidates[0]
    return None


def find_ncx(folder: str) -> str | None:
    for root, _, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".ncx"):
                return os.path.join(root, fn)
    return None


def find_opf(folder: str) -> str | None:
    for root, _, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".opf"):
                return os.path.join(root, fn)
    return None


def convert_zip_to_text(zip_path: str) -> tuple[str, str] | None:
    """ZIP 으로부터 압축 해제 + 본문 TXT 생성을 한 번에 처리.

    반환: (추출 폴더, TXT 파일 경로). 실패 시 None.
    """
    if not is_daisy_zip(zip_path):
        return None
    folder = extract_zip(zip_path)
    xml = find_dtbook_xml(folder)
    if not xml:
        return None
    body_text = dtbook_to_text(xml)

    # 제목·목차 추가
    opf = find_opf(folder)
    ncx = find_ncx(folder)
    title = ""
    for p in (opf, ncx):
        if p:
            title = book_title(p)
            if title:
                break
    toc = ncx_to_toc(ncx) if ncx else ""

    parts: list[str] = []
    if title:
        parts.append(f"# {title}\n")
    if toc:
        parts.append("## 목차\n")
        parts.append(toc + "\n")
    parts.append("## 본문\n")
    parts.append(body_text)

    base = os.path.splitext(os.path.basename(zip_path))[0]
    txt_path = os.path.join(folder, f"{base}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return folder, txt_path
