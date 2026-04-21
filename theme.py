"""초록멀티 테마 - 저시력 사용자 최적화 색상/글꼴"""
import os
import wx

from config import DATA_DIR


# ── 테마 프리셋 (저시력자 선호 조합) ──
# 각 테마: (이름, 배경RGB, 글자RGB, 상태바배경RGB, 상태바글자RGB,
#            버튼배경RGB, 버튼글자RGB, 입력창배경RGB, 입력창글자RGB)

THEME_PRESETS = {
    "dark_yellow": {
        "name": "검정 바탕 + 노란 글씨",
        "bg": (0, 0, 0),
        "fg": (255, 255, 0),
        "status_bg": (30, 30, 0),
        "status_fg": (255, 230, 80),
        "btn_bg": (50, 50, 10),
        "btn_fg": (255, 255, 0),
        "input_bg": (15, 15, 0),
        "input_fg": (255, 255, 100),
    },
    "dark_white": {
        "name": "검정 바탕 + 하얀 글씨",
        "bg": (0, 0, 0),
        "fg": (255, 255, 255),
        "status_bg": (30, 30, 30),
        "status_fg": (220, 220, 220),
        "btn_bg": (50, 50, 50),
        "btn_fg": (255, 255, 255),
        "input_bg": (15, 15, 15),
        "input_fg": (240, 240, 240),
    },
    "dark_green": {
        "name": "검정 바탕 + 초록 글씨",
        "bg": (0, 0, 0),
        "fg": (0, 255, 100),
        "status_bg": (0, 30, 10),
        "status_fg": (100, 255, 150),
        "btn_bg": (0, 50, 20),
        "btn_fg": (0, 255, 100),
        "input_bg": (0, 15, 5),
        "input_fg": (100, 255, 160),
    },
    "navy_yellow": {
        "name": "남색 바탕 + 노란 글씨",
        "bg": (26, 31, 48),
        "fg": (255, 241, 118),
        "status_bg": (15, 22, 38),
        "status_fg": (255, 224, 102),
        "btn_bg": (52, 73, 110),
        "btn_fg": (255, 241, 118),
        "input_bg": (10, 14, 24),
        "input_fg": (255, 245, 157),
    },
    "blue_white": {
        "name": "파란 바탕 + 하얀 글씨",
        "bg": (0, 0, 128),
        "fg": (255, 255, 255),
        "status_bg": (0, 0, 90),
        "status_fg": (200, 220, 255),
        "btn_bg": (30, 30, 160),
        "btn_fg": (255, 255, 255),
        "input_bg": (0, 0, 100),
        "input_fg": (240, 240, 255),
    },
    "white_black": {
        "name": "하얀 바탕 + 검정 글씨",
        "bg": (255, 255, 255),
        "fg": (0, 0, 0),
        "status_bg": (230, 230, 230),
        "status_fg": (30, 30, 30),
        "btn_bg": (220, 220, 220),
        "btn_fg": (0, 0, 0),
        "input_bg": (245, 245, 245),
        "input_fg": (0, 0, 0),
    },
    "cream_navy": {
        "name": "아이보리 바탕 + 남색 글씨",
        "bg": (255, 253, 230),
        "fg": (0, 0, 80),
        "status_bg": (240, 238, 215),
        "status_fg": (0, 0, 60),
        "btn_bg": (230, 228, 210),
        "btn_fg": (0, 0, 80),
        "input_bg": (250, 248, 225),
        "input_fg": (0, 0, 60),
    },
    "yellow_black": {
        "name": "노란 바탕 + 검정 글씨",
        "bg": (255, 255, 180),
        "fg": (0, 0, 0),
        "status_bg": (240, 240, 160),
        "status_fg": (20, 20, 0),
        "btn_bg": (230, 230, 150),
        "btn_fg": (0, 0, 0),
        "input_bg": (250, 250, 175),
        "input_fg": (0, 0, 0),
    },
}

# 테마 키 순서 (대화상자 목록 순서)
THEME_ORDER = [
    "white_black",
    "dark_yellow",
    "dark_white",
    "dark_green",
    "navy_yellow",
    "blue_white",
    "cream_navy",
    "yellow_black",
]

DEFAULT_THEME_KEY = "white_black"

# ── 현재 테마 색상 (모듈 변수, set_current_theme으로 변경) ──
COLOR_BG_MAIN = wx.Colour(26, 31, 48)
COLOR_FG_MAIN = wx.Colour(255, 241, 118)
COLOR_BG_STATUS = wx.Colour(15, 22, 38)
COLOR_FG_STATUS = wx.Colour(255, 224, 102)
COLOR_BG_BUTTON = wx.Colour(52, 73, 110)
COLOR_FG_BUTTON = wx.Colour(255, 241, 118)
COLOR_BG_INPUT = wx.Colour(10, 14, 24)
COLOR_FG_INPUT = wx.Colour(255, 245, 157)


# ── 글꼴 ──
DEFAULT_FONT_SIZE = 12
MIN_FONT_SIZE = 9
MAX_FONT_SIZE = 40
FONT_SIZE_STEP = 2
FONT_FACE = "맑은 고딕"

# ── 저장 파일 ──
FONT_SIZE_FILE = os.path.join(DATA_DIR, "font_size.txt")
THEME_FILE = os.path.join(DATA_DIR, "theme.txt")


def _update_module_colors(preset: dict) -> None:
    """모듈 변수에 현재 테마 색상을 반영한다."""
    global COLOR_BG_MAIN, COLOR_FG_MAIN
    global COLOR_BG_STATUS, COLOR_FG_STATUS
    global COLOR_BG_BUTTON, COLOR_FG_BUTTON
    global COLOR_BG_INPUT, COLOR_FG_INPUT

    COLOR_BG_MAIN = wx.Colour(*preset["bg"])
    COLOR_FG_MAIN = wx.Colour(*preset["fg"])
    COLOR_BG_STATUS = wx.Colour(*preset["status_bg"])
    COLOR_FG_STATUS = wx.Colour(*preset["status_fg"])
    COLOR_BG_BUTTON = wx.Colour(*preset["btn_bg"])
    COLOR_FG_BUTTON = wx.Colour(*preset["btn_fg"])
    COLOR_BG_INPUT = wx.Colour(*preset["input_bg"])
    COLOR_FG_INPUT = wx.Colour(*preset["input_fg"])


def load_theme_key() -> str:
    """저장된 테마 키를 불러온다."""
    try:
        if os.path.exists(THEME_FILE):
            with open(THEME_FILE, "r", encoding="utf-8") as f:
                key = f.read().strip()
                if key in THEME_PRESETS:
                    return key
    except Exception:
        pass
    return DEFAULT_THEME_KEY


def save_theme_key(key: str) -> None:
    """테마 키를 저장한다."""
    try:
        os.makedirs(os.path.dirname(THEME_FILE), exist_ok=True)
        with open(THEME_FILE, "w", encoding="utf-8") as f:
            f.write(key)
    except Exception:
        pass


def set_current_theme(key: str) -> None:
    """테마를 변경한다 (모듈 색상 업데이트 + 파일 저장)."""
    preset = THEME_PRESETS.get(key, THEME_PRESETS[DEFAULT_THEME_KEY])
    _update_module_colors(preset)
    save_theme_key(key)


def get_current_theme_name() -> str:
    """현재 테마의 이름을 반환한다."""
    key = load_theme_key()
    return THEME_PRESETS.get(key, THEME_PRESETS[DEFAULT_THEME_KEY])["name"]


def init_theme() -> None:
    """프로그램 시작 시 저장된 테마를 로드하여 모듈 색상을 설정한다."""
    key = load_theme_key()
    preset = THEME_PRESETS.get(key, THEME_PRESETS[DEFAULT_THEME_KEY])
    _update_module_colors(preset)


# ── 글꼴 저장/불러오기 ──

def load_font_size() -> int:
    """저장된 글꼴 크기를 불러온다."""
    try:
        if os.path.exists(FONT_SIZE_FILE):
            with open(FONT_SIZE_FILE, "r", encoding="utf-8") as f:
                size = int(f.read().strip())
                if MIN_FONT_SIZE <= size <= MAX_FONT_SIZE:
                    return size
    except Exception:
        pass
    return DEFAULT_FONT_SIZE


def save_font_size(size: int) -> None:
    """글꼴 크기를 저장한다."""
    try:
        os.makedirs(os.path.dirname(FONT_SIZE_FILE), exist_ok=True)
        with open(FONT_SIZE_FILE, "w", encoding="utf-8") as f:
            f.write(str(size))
    except Exception:
        pass


def make_font(size: int, bold: bool = True) -> wx.Font:
    """주어진 크기의 테마 글꼴을 생성한다."""
    weight = wx.FONTWEIGHT_BOLD if bold else wx.FONTWEIGHT_NORMAL
    return wx.Font(
        size,
        wx.FONTFAMILY_DEFAULT,
        wx.FONTSTYLE_NORMAL,
        weight,
        faceName=FONT_FACE,
    )


def apply_theme(widget: wx.Window, font: wx.Font | None = None) -> None:
    """위젯과 모든 자식에 테마를 재귀 적용한다."""
    _apply_colors(widget)
    if font is not None:
        _apply_font(widget, font)
    widget.Refresh()


def _apply_colors(widget: wx.Window) -> None:
    """위젯 타입별 색상 적용 후 자식에 재귀."""
    if isinstance(widget, wx.TextCtrl):
        widget.SetBackgroundColour(COLOR_BG_INPUT)
        widget.SetForegroundColour(COLOR_FG_INPUT)
    elif isinstance(widget, (wx.ListBox, wx.ComboBox, wx.Choice)):
        widget.SetBackgroundColour(COLOR_BG_INPUT)
        widget.SetForegroundColour(COLOR_FG_INPUT)
    elif isinstance(widget, wx.CheckBox):
        pass  # 네이티브 확인란 역할/렌더링 유지를 위해 색상 적용하지 않음
    elif isinstance(widget, wx.Button):
        widget.SetBackgroundColour(COLOR_BG_BUTTON)
        widget.SetForegroundColour(COLOR_FG_BUTTON)
    elif isinstance(widget, wx.StatusBar):
        widget.SetBackgroundColour(COLOR_BG_STATUS)
        widget.SetForegroundColour(COLOR_FG_STATUS)
    elif isinstance(widget, (wx.Panel, wx.Dialog, wx.Frame)):
        widget.SetBackgroundColour(COLOR_BG_MAIN)
        widget.SetForegroundColour(COLOR_FG_MAIN)
    else:
        widget.SetBackgroundColour(COLOR_BG_MAIN)
        widget.SetForegroundColour(COLOR_FG_MAIN)

    for child in widget.GetChildren():
        _apply_colors(child)


def _apply_font(widget: wx.Window, font: wx.Font) -> None:
    """위젯과 자식에 글꼴 적용."""
    widget.SetFont(font)
    for child in widget.GetChildren():
        _apply_font(child, font)


# 모듈 로드 시 저장된 테마 자동 적용
init_theme()
