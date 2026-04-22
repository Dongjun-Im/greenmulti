"""사운드 이벤트 재생 및 사용자 설정 저장/불러오기.

- 이벤트별 WAV 경로를 사용자 지정 가능
- 사운드 마스터 on/off 토글
- 설정은 data/sound_settings.json 에 저장
- 사용자 지정 경로가 없으면 sounds/<event_key>.wav (번들) 사용
"""
import json
import os
import platform

from config import DATA_DIR, SOUNDS_DIR


SOUND_SETTINGS_FILE = os.path.join(DATA_DIR, "sound_settings.json")


# 사용자에게 표시할 이벤트 목록. 순서도 설정 UI에 그대로 반영.
SOUND_EVENTS: list[tuple[str, str]] = [
    ("program_start", "프로그램 시작"),
    ("program_end", "프로그램 종료"),
    ("page_move", "게시물 페이지 이동"),
    ("main_menu_return", "메인메뉴로 돌아왔을 때"),
    ("home_end", "홈/엔드 키 경고음"),
    ("download_start", "파일 다운로드 시작"),
    ("download_complete", "파일 다운로드 완료"),
]

EVENT_KEYS = [k for k, _ in SOUND_EVENTS]


def _default_settings() -> dict:
    return {
        "enabled": True,
        "events": {k: "" for k in EVENT_KEYS},
        "event_enabled": {k: True for k in EVENT_KEYS},
    }


def load_sound_settings() -> dict:
    """저장된 사운드 설정을 불러온다. 누락된 항목은 기본값으로 보강.
    이전 포맷(events만 있음)도 그대로 읽을 수 있도록 방어적으로 처리."""
    settings = _default_settings()
    try:
        if os.path.exists(SOUND_SETTINGS_FILE):
            with open(SOUND_SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                settings["enabled"] = bool(data.get("enabled", True))
                ev = data.get("events") or {}
                if isinstance(ev, dict):
                    for k in EVENT_KEYS:
                        v = ev.get(k, "")
                        if isinstance(v, str):
                            settings["events"][k] = v
                een = data.get("event_enabled") or {}
                if isinstance(een, dict):
                    for k in EVENT_KEYS:
                        if k in een:
                            settings["event_enabled"][k] = bool(een[k])
    except Exception:
        pass
    return settings


def save_sound_settings(settings: dict) -> None:
    """사운드 설정을 저장."""
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        cleaned = {
            "enabled": bool(settings.get("enabled", True)),
            "events": {
                k: str((settings.get("events") or {}).get(k, ""))
                for k in EVENT_KEYS
            },
            "event_enabled": {
                k: bool((settings.get("event_enabled") or {}).get(k, True))
                for k in EVENT_KEYS
            },
        }
        with open(SOUND_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def resolve_event_path(event_key: str, settings: dict | None = None) -> str:
    """이벤트에 대응하는 WAV 파일 경로를 결정.
    1) 사용자 지정 경로가 유효하면 그 경로
    2) 아니면 sounds/<event_key>.wav (번들 기본)
    3) 둘 다 없으면 빈 문자열
    """
    if settings is None:
        settings = load_sound_settings()
    custom = (settings.get("events") or {}).get(event_key, "")
    if custom and os.path.exists(custom):
        return custom
    default = os.path.join(SOUNDS_DIR, f"{event_key}.wav")
    if os.path.exists(default):
        return default
    return ""


def play_event(event_key: str, block: bool = False) -> bool:
    """이벤트 사운드 재생. 마스터 스위치가 꺼져 있거나 해당 이벤트가 꺼져 있거나
    파일이 없으면 조용히 무시한다. block=True면 동기 재생(종료 사운드용).
    반환값: 실제로 재생 시도됐는지."""
    if platform.system() != "Windows":
        return False
    try:
        settings = load_sound_settings()
        if not settings.get("enabled", True):
            return False
        event_enabled = (settings.get("event_enabled") or {}).get(event_key, True)
        if not event_enabled:
            return False
        path = resolve_event_path(event_key, settings)
        if not path:
            return False
        import winsound
        flags = winsound.SND_FILENAME
        if not block:
            flags |= winsound.SND_ASYNC
        winsound.PlaySound(path, flags)
        return True
    except Exception:
        return False


def play_file(path: str, block: bool = False) -> bool:
    """임의 WAV 파일을 직접 재생 (설정 대화상자의 '듣기' 버튼용)."""
    if platform.system() != "Windows":
        return False
    if not path or not os.path.exists(path):
        return False
    try:
        import winsound
        flags = winsound.SND_FILENAME
        if not block:
            flags |= winsound.SND_ASYNC
        winsound.PlaySound(path, flags)
        return True
    except Exception:
        return False
