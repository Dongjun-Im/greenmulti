"""스크린리더 음성 출력 모듈 (NVDA, 센스리더 지원)"""
import ctypes
import platform


def speak(text: str) -> bool:
    """
    스크린리더로 텍스트를 음성 출력한다.
    NVDA -> 센스리더 순서로 시도하며, 둘 다 없으면 False를 반환한다.
    """
    if _speak_nvda(text):
        return True
    if _speak_sense_reader(text):
        return True
    return False


def _speak_nvda(text: str) -> bool:
    """NVDA 스크린리더로 음성 출력"""
    if platform.system() != "Windows":
        return False
    try:
        nvda_dll = ctypes.windll.LoadLibrary("nvdaControllerClient64.dll")
    except OSError:
        try:
            nvda_dll = ctypes.windll.LoadLibrary("nvdaControllerClient32.dll")
        except OSError:
            return False

    try:
        result = nvda_dll.nvdaController_testIfRunning()
        if result != 0:
            return False
        nvda_dll.nvdaController_cancelSpeech()
        nvda_dll.nvdaController_speakText(text)
        return True
    except Exception:
        return False


def _speak_sense_reader(text: str) -> bool:
    """센스리더 COM 자동화로 음성 출력"""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client
        app = win32com.client.Dispatch("SenseReader.Application")
        app.Speak(text)
        return True
    except Exception:
        pass

    try:
        import comtypes.client
        app = comtypes.client.CreateObject("SenseReader.Application")
        app.Speak(text)
        return True
    except Exception:
        return False
