"""알림 센터 — 쪽지·메일 등 여러 종류의 알림을 통합 관리.

특징:
- 알림 추가/삭제/전체 삭제
- 중복 방지 (같은 type + item_id 는 다시 추가 안 됨)
- 스레드 안전 (lock)
- 외부에서 관찰할 수 있는 변경 콜백 (UI 갱신용)
"""
from __future__ import annotations

import threading
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Optional


@dataclass
class NotificationItem:
    """알림 항목 하나.

    type: "memo" 또는 "mail"
    item_id: 타입별 고유 ID (쪽지의 me_id, 메일의 message_id 등)
    sender: 표시용 보낸 사람 이름/아이디
    summary: 제목 또는 요약
    timestamp: 서버 표기 시각(원문) 또는 클라이언트 감지 시각
    received_at: 클라이언트에서 감지한 시각 (정렬용)
    extra: 원본 MemoItem/MailItem 등 참조 보관 (열람 시 사용)
    """
    type: str
    item_id: str
    sender: str
    summary: str
    timestamp: str = ""
    received_at: datetime = field(default_factory=datetime.now)
    extra: Any = None


class NotificationCenter:
    """알림 저장소 + 변경 리스너 관리."""

    def __init__(self):
        self._lock = threading.RLock()
        self._items: list[NotificationItem] = []
        self._listeners: list[Callable[[], None]] = []

    # ── 리스너 ──

    def add_listener(self, cb: Callable[[], None]):
        with self._lock:
            if cb not in self._listeners:
                self._listeners.append(cb)

    def remove_listener(self, cb: Callable[[], None]):
        with self._lock:
            if cb in self._listeners:
                self._listeners.remove(cb)

    def _notify(self):
        # 리스너 호출은 lock 밖에서
        with self._lock:
            listeners = list(self._listeners)
        for cb in listeners:
            try:
                cb()
            except Exception:
                pass

    # ── CRUD ──

    def add(self, item: NotificationItem) -> bool:
        """중복(동일 type+item_id) 이면 False, 새로 추가되면 True."""
        with self._lock:
            for existing in self._items:
                if existing.type == item.type and existing.item_id == item.item_id:
                    return False
            self._items.append(item)
        self._notify()
        return True

    def add_many(self, items: list[NotificationItem]) -> int:
        """여러 개 일괄 추가. 추가된 개수 반환 (중복 제외)."""
        added = 0
        with self._lock:
            existing_keys = {(it.type, it.item_id) for it in self._items}
            for item in items:
                key = (item.type, item.item_id)
                if key in existing_keys:
                    continue
                self._items.append(item)
                existing_keys.add(key)
                added += 1
        if added > 0:
            self._notify()
        return added

    def remove(self, type_: str, item_id: str) -> bool:
        with self._lock:
            for i, existing in enumerate(self._items):
                if existing.type == type_ and existing.item_id == item_id:
                    del self._items[i]
                    break
            else:
                return False
        self._notify()
        return True

    def remove_at(self, index: int) -> Optional[NotificationItem]:
        with self._lock:
            if 0 <= index < len(self._items):
                removed = self._items.pop(index)
            else:
                return None
        self._notify()
        return removed

    def clear_all(self) -> int:
        """모든 알림 삭제. 삭제한 개수 반환."""
        with self._lock:
            count = len(self._items)
            self._items = []
        if count > 0:
            self._notify()
        return count

    def items(self) -> list[NotificationItem]:
        """알림 목록의 스냅샷(최신 순)."""
        with self._lock:
            return list(self._items)

    def count(self) -> int:
        with self._lock:
            return len(self._items)

    def count_by_type(self, type_: str) -> int:
        with self._lock:
            return sum(1 for it in self._items if it.type == type_)


# 전역 싱글톤
_center: Optional[NotificationCenter] = None


def get_center() -> NotificationCenter:
    global _center
    if _center is None:
        _center = NotificationCenter()
    return _center
