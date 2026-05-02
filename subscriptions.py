"""게시판 새 글 알림 / 구독 관리 (v1.7).

사용자가 관심 게시판을 구독하면, MemoNotifier·MailNotifier 와 같은 방식으로
백그라운드에서 주기적으로 게시판 글 목록을 조회해 새 글을 발견하면 알림
센터에 등록한다.

저장 경로: data/subscriptions.json
"""

from __future__ import annotations

import json
import os
import threading
from dataclasses import dataclass, asdict, field
from typing import Callable

import requests

from config import SUBSCRIPTIONS_FILE, SORISEM_BASE_URL
from page_parser import parse_board_list


@dataclass
class Subscription:
    name: str
    url: str
    seen_ids: set[str] = field(default_factory=set)


class SubscriptionManager:
    """구독 목록 + 폴링."""

    def __init__(self, parent_frame, session: requests.Session,
                 callback_on_new):
        """
        callback_on_new(subscription_name, new_post_items): 새 글이 있을 때
        UI 스레드에서 호출되는 콜백 (PostItem 리스트).
        """
        self.frame = parent_frame
        self.session = session
        self.callback = callback_on_new
        self.items: list[Subscription] = []
        self._initial_done = False
        self._in_flight = False
        self.load()

    # ── 영속성 ──

    def load(self) -> None:
        if not os.path.exists(SUBSCRIPTIONS_FILE):
            self.items = []
            return
        try:
            with open(SUBSCRIPTIONS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            self.items = []
            return
        out: list[Subscription] = []
        for it in data.get("items", []) or []:
            try:
                sub = Subscription(
                    name=str(it.get("name", "")).strip(),
                    url=str(it.get("url", "")).strip(),
                    seen_ids=set(it.get("seen_ids", []) or []),
                )
            except Exception:
                continue
            if sub.name and sub.url:
                out.append(sub)
        self.items = out

    def save(self) -> None:
        os.makedirs(os.path.dirname(SUBSCRIPTIONS_FILE), exist_ok=True)
        data = {
            "version": 1,
            "items": [
                {"name": s.name, "url": s.url, "seen_ids": list(s.seen_ids)}
                for s in self.items
            ],
        }
        with open(SUBSCRIPTIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    # ── 변경 ──

    def find(self, url: str) -> Subscription | None:
        for s in self.items:
            if s.url == url:
                return s
        return None

    def add(self, name: str, url: str) -> bool:
        if self.find(url):
            return False
        self.items.append(Subscription(name=name, url=url, seen_ids=set()))
        self.save()
        return True

    def remove(self, url: str) -> bool:
        before = len(self.items)
        self.items = [s for s in self.items if s.url != url]
        if len(self.items) != before:
            self.save()
            return True
        return False

    # ── 폴링 ──

    def initial_fill_async(self) -> None:
        """시작 시 각 구독 게시판의 현재 글 ID 들을 seen 으로 등록.

        seen 이 비어 있는 구독에 대해서만 처음 한 번 채운다 (이후 폴링에서
        새 글 감지). 이미 seen 이 채워져 있는 구독은 그대로 둔다.
        """
        targets = [s for s in self.items if not s.seen_ids]

        def worker():
            try:
                for sub in targets:
                    posts = self._fetch_board_posts(sub.url)
                    for p in posts:
                        pid = self._post_id(p)
                        if pid:
                            sub.seen_ids.add(pid)
                if targets:
                    self.save()
            finally:
                self._initial_done = True

        if not targets:
            self._initial_done = True
            return
        threading.Thread(target=worker, daemon=True).start()

    def poll_once_async(self) -> None:
        """1회 폴링 — 모든 구독에 대해 새 글 검사. 백그라운드 스레드."""
        if self._in_flight:
            return
        if not self._initial_done:
            return
        if not self.items:
            return
        self._in_flight = True

        def worker():
            try:
                for sub in self.items:
                    try:
                        posts = self._fetch_board_posts(sub.url)
                    except Exception:
                        continue
                    new_posts = []
                    for p in posts:
                        pid = self._post_id(p)
                        if pid and pid not in sub.seen_ids:
                            new_posts.append(p)
                            sub.seen_ids.add(pid)
                    if new_posts:
                        try:
                            import wx as _wx
                            _wx.CallAfter(self.callback, sub.name, new_posts)
                        except Exception:
                            pass
                self.save()
            finally:
                self._in_flight = False

        threading.Thread(target=worker, daemon=True).start()

    # ── 내부 ──

    def _fetch_board_posts(self, url: str) -> list:
        full = url if url.startswith("http") else f"{SORISEM_BASE_URL}{url}"
        resp = self.session.get(full, timeout=15)
        if resp.status_code != 200:
            return []
        return parse_board_list(resp.text) or []

    @staticmethod
    def _post_id(post) -> str:
        """PostItem 에서 안정적으로 식별자 추출. wr_id 우선, 없으면 url 자체."""
        import re
        url = getattr(post, "url", "") or ""
        m = re.search(r"wr_id=(\d+)", url)
        if m:
            return m.group(1)
        return url
