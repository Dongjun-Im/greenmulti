"""즐겨찾기(Bookmark) 관리 (v1.7).

자주 가는 게시판·게시물을 즐겨찾기로 등록해 두고 빠르게 다시 진입한다.
저장 경로: data/bookmarks.json
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict
from typing import Iterable

from config import BOOKMARKS_FILE


@dataclass
class Bookmark:
    """단일 즐겨찾기 항목."""
    name: str
    url: str
    type: str = "board"  # board / post / category 등 — 표시용 힌트


class BookmarkManager:
    """즐겨찾기 목록 로드·저장·추가·삭제."""

    def __init__(self):
        self.items: list[Bookmark] = []
        self.load()

    def load(self) -> None:
        if not os.path.exists(BOOKMARKS_FILE):
            self.items = []
            return
        try:
            with open(BOOKMARKS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            self.items = []
            return
        out: list[Bookmark] = []
        for it in data.get("items", []) or []:
            try:
                out.append(Bookmark(
                    name=str(it.get("name", "")).strip(),
                    url=str(it.get("url", "")).strip(),
                    type=str(it.get("type", "board")),
                ))
            except Exception:
                continue
        self.items = [b for b in out if b.name and b.url]

    def save(self) -> None:
        os.makedirs(os.path.dirname(BOOKMARKS_FILE), exist_ok=True)
        data = {
            "version": 1,
            "description": "초록멀티 즐겨찾기",
            "items": [asdict(b) for b in self.items],
        }
        with open(BOOKMARKS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def add(self, name: str, url: str, type_hint: str = "board") -> bool:
        """이미 있는 URL 이면 이름만 갱신. 새 URL 이면 추가."""
        name = (name or "").strip()
        url = (url or "").strip()
        if not name or not url:
            return False
        for b in self.items:
            if b.url == url:
                b.name = name
                b.type = type_hint
                self.save()
                return False  # 중복(이름만 갱신)
        self.items.append(Bookmark(name=name, url=url, type=type_hint))
        self.save()
        return True

    def remove(self, index: int) -> bool:
        if 0 <= index < len(self.items):
            del self.items[index]
            self.save()
            return True
        return False

    def reorder(self, src: int, dst: int) -> bool:
        if not (0 <= src < len(self.items)) or not (0 <= dst < len(self.items)):
            return False
        item = self.items.pop(src)
        self.items.insert(dst, item)
        self.save()
        return True

    def __len__(self) -> int:
        return len(self.items)

    def __iter__(self) -> Iterable[Bookmark]:
        return iter(self.items)

    def get(self, index: int) -> Bookmark | None:
        if 0 <= index < len(self.items):
            return self.items[index]
        return None
