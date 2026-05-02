"""빠른 답장 템플릿 관리 (v1.7).

사용자가 자주 쓰는 댓글·답장 문구를 한 줄씩 저장해 두고, 댓글·답장 입력
대화상자에서 Alt+숫자(1~9) 로 커서 위치에 즉시 삽입한다.

파일 포맷 — `data/reply_templates.txt`:
    한 줄당 한 템플릿. '#' 으로 시작하는 줄은 주석, 빈 줄은 무시.
    파일이 없으면 적당한 기본값으로 자동 생성한다.
"""

from __future__ import annotations

import os

from config import REPLY_TEMPLATES_FILE


_DEFAULT_TEMPLATES = [
    "감사합니다.",
    "좋은 글 감사합니다.",
    "잘 읽었습니다.",
    "확인했습니다.",
    "좋은 정보 감사합니다.",
    "동감합니다.",
    "고생 많으셨습니다.",
    "축하드립니다.",
    "응원합니다.",
]


_HEADER = """\
# 빠른 답장 템플릿 — 한 줄에 하나씩
# 댓글 / 답장 / 메일 입력 창에서 Alt+1 ~ Alt+9 로 커서 위치에 삽입됩니다.
# '#' 으로 시작하는 줄과 빈 줄은 무시됩니다.
"""


def load_templates() -> list[str]:
    """저장된 템플릿 목록을 한 줄당 한 개씩 반환. 최대 9개.

    파일이 없으면 기본 템플릿으로 자동 생성한다.
    """
    if not os.path.exists(REPLY_TEMPLATES_FILE):
        try:
            save_templates(_DEFAULT_TEMPLATES)
        except OSError:
            pass
        return list(_DEFAULT_TEMPLATES)[:9]

    try:
        with open(REPLY_TEMPLATES_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except OSError:
        return list(_DEFAULT_TEMPLATES)[:9]

    items: list[str] = []
    for raw in lines:
        line = raw.rstrip("\n").rstrip("\r")
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        items.append(line.strip())
        if len(items) >= 9:
            break
    return items


def save_templates(items: list[str]) -> None:
    """템플릿 목록을 파일에 저장 (헤더 주석 + 한 줄당 하나)."""
    os.makedirs(os.path.dirname(REPLY_TEMPLATES_FILE), exist_ok=True)
    text = _HEADER + "\n" + "\n".join(items) + "\n"
    with open(REPLY_TEMPLATES_FILE, "w", encoding="utf-8") as f:
        f.write(text)


def get_template(index: int) -> str:
    """1..9 번 템플릿. 없으면 빈 문자열."""
    items = load_templates()
    if 1 <= index <= len(items):
        return items[index - 1]
    return ""
