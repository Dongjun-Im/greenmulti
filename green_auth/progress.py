"""인증 진행 중 비프음 + 스크린리더 진행 알림"""
import threading
import time

from green_auth.screen_reader import speak


class ProgressIndicator:
    """별도 스레드에서 비프음을 주기적으로 재생하면서 진행 상황을 알린다.

    비프음은 음높이를 단계별로 변화시켜 진행감(프로그레스)을 표현하고,
    일정 간격마다 스크린리더에도 짧은 진행 알림을 출력한다.
    """

    BEEP_INTERVAL = 0.5
    BEEP_DURATION_MS = 80
    BEEP_BASE_FREQ = 700
    BEEP_STEP_FREQ = 80
    BEEP_STEPS = 5
    SPEAK_INTERVAL = 3.0
    # 시작 후 INITIAL_SPEAK_DELAY 초가 지나기 전까지는 진행 알림 발화를 보류.
    # 초기 "인증 중입니다. 잠시만 기다려 주세요." 발화가 잘리지 않게 보호.
    INITIAL_SPEAK_DELAY = 6.0

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        if self._thread is None:
            return
        self._stop_event.set()
        self._thread.join(timeout=1.5)
        self._thread = None

    def _run(self) -> None:
        try:
            import winsound
        except ImportError:
            winsound = None

        start = time.monotonic()
        last_speak = start
        step = 0
        while not self._stop_event.wait(self.BEEP_INTERVAL):
            if winsound is not None:
                freq = self.BEEP_BASE_FREQ + (step % self.BEEP_STEPS) * self.BEEP_STEP_FREQ
                try:
                    winsound.Beep(freq, self.BEEP_DURATION_MS)
                except Exception:
                    pass
            step += 1
            now = time.monotonic()
            # 시작 직후 N초 동안은 발화하지 않음 — 초기 "인증 중입니다" 발화 보호.
            if now - start < self.INITIAL_SPEAK_DELAY:
                continue
            if now - last_speak >= self.SPEAK_INTERVAL:
                speak("인증 진행 중")
                last_speak = now
