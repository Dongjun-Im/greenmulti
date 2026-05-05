"""스크린리더 음성 출력 모듈.

지원 대상:
- NVDA (NV Access)
- 보이스위드 (NVDA 기반 한국어 스크린리더 — VoiceWith 가 동봉한 32-bit DLL
  대신 accessible_output2 가 들고 있는 64-bit nvdaControllerClient64.dll 로
  통신해 64-bit 앱에서도 정상 작동)
- 센스리더 (엑스비전 — COM 자동화로 호출)
- Windows SAPI 5 (위 셋 모두 없을 때 폴백)

발화 우선순위:
1. accessible_output2.outputs.auto.Auto — NVDA / JAWS / SAPI 등 자동 감지
   (라이브러리에 64-bit nvdaControllerClient64.dll 이 동봉되어 있어 NVDA 와
   보이스위드 모두 64-bit 프로세스에서 안정적으로 호출 가능)
2. nvdaControllerClient* 직접 호출 (DLL 다중 경로 탐색 폴백)
3. 센스리더 COM 자동화
4. SAPI 5

DLL 핸들과 COM 객체는 import 시점에 한 번 시도해 캐싱한다.
"""
from __future__ import annotations

import ctypes
import os
import platform
import subprocess
import sys
import threading
from typing import Optional


# ─────────────── 보이스위드 32-bit DLL 프록시 (PowerShell helper) ───────────────

# 보이스위드는 NVDA 의 fork 지만 자체 RPC 서버를 쓰는 듯, 64-bit
# nvdaControllerClient64.dll 로는 발화가 닿지 않는다. 보이스위드가 동봉한
# 32-bit `nvdaControllerClient.dll` 만 보이스위드와 통신할 수 있는데, 우리 앱은
# 64-bit 라 ctypes 로 직접 로드 불가. → Windows 가 제공하는 32-bit PowerShell
# (SysWOW64) 을 장수명 helper 프로세스로 띄우고 stdin 으로 발화 텍스트를 보내,
# 그 32-bit 프로세스 안에서 P/Invoke 로 보이스위드 DLL 을 호출하도록 우회한다.

_VOICEWITH_DLL_CANDIDATES = [
    r"C:\Program Files (x86)\VoiceWith\nvdaControllerClient.dll",
    r"C:\Program Files\VoiceWith\nvdaControllerClient.dll",
]
_POWERSHELL_32_PATHS = [
    r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
]

_VW_PROXY: Optional[subprocess.Popen] = None
_VW_LOCK = threading.Lock()
# 첫 발화는 RC 응답을 동기 검증 (보이스위드 동작 여부 판정).
# 두 번째 발화부터는 fire-and-forget — stdout 은 백그라운드 스레드가 드레인.
_VW_VERIFIED = False
_VW_DRAIN_THREAD: Optional[threading.Thread] = None
_VW_DISABLED = False  # 첫 발화 실패 시 True 로 설정해 이후 시도 안 함


def _resolve_voicewith_dll() -> Optional[str]:
    for p in _VOICEWITH_DLL_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


def _resolve_powershell_32() -> Optional[str]:
    for p in _POWERSHELL_32_PATHS:
        if os.path.isfile(p):
            return p
    return None


def _spawn_voicewith_proxy() -> Optional[subprocess.Popen]:
    """SysWOW64 PowerShell 을 장수명 helper 로 띄워 보이스위드 DLL 을 로드.

    프로토콜:
      - stdin 한 줄 = 발화할 텍스트 (UTF-8). 빈 줄/SPC 만 있는 줄은 무시.
      - 특수 명령 "__cancel__" 한 줄 = nvdaController_cancelSpeech()
      - stdout 한 줄 = 마지막 작업의 반환 코드 ("RC:0" 등). 호출자는 옵션으로 읽음.
    """
    dll = _resolve_voicewith_dll()
    ps = _resolve_powershell_32()
    if dll is None or ps is None:
        return None

    # 닫는 따옴표·중괄호 이스케이프 주의 — Add-Type 은 single-quoted here-string 사용.
    # P/Invoke 시 CharSet.Unicode 로 wide-char 진입점에 맞춤.
    # 콘솔 입출력 인코딩을 UTF-8 로 강제 — 그러지 않으면 cp949(기본 Windows
    # 콘솔 코드페이지) 로 디코딩되어 한글이 깨진다.
    # [Console]::InputEncoding 은 stdin 읽기 인코딩, OutputEncoding 은 stdout 쓰기.
    # `$OutputEncoding` 도 함께 설정해야 PowerShell 의 native 명령 출력이 UTF-8.
    script = (
        "[Console]::InputEncoding = "
        "New-Object System.Text.UTF8Encoding $false\n"
        "[Console]::OutputEncoding = "
        "New-Object System.Text.UTF8Encoding $false\n"
        "$OutputEncoding = "
        "New-Object System.Text.UTF8Encoding $false\n"
        "Add-Type -TypeDefinition @'\n"
        "using System;\n"
        "using System.Runtime.InteropServices;\n"
        "public static class VW {\n"
        f'  [DllImport(@"{dll}", CharSet = CharSet.Unicode)]\n'
        "  public static extern int nvdaController_speakText(string text);\n"
        f'  [DllImport(@"{dll}")]\n'
        "  public static extern int nvdaController_cancelSpeech();\n"
        "}\n"
        "'@\n"
        "while ($line = [System.Console]::In.ReadLine()) {\n"
        '  if ($line -eq "__cancel__") {\n'
        "    $rc = [VW]::nvdaController_cancelSpeech()\n"
        "  } else {\n"
        "    [void][VW]::nvdaController_cancelSpeech()\n"
        "    $rc = [VW]::nvdaController_speakText($line)\n"
        "  }\n"
        '  [System.Console]::Out.WriteLine("RC:" + $rc)\n'
        "  [System.Console]::Out.Flush()\n"
        "}\n"
    )

    try:
        # CREATE_NO_WINDOW=0x08000000 — 콘솔 창이 깜빡이지 않게.
        flags = 0x08000000 if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        proc = subprocess.Popen(
            [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-Command", script],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            creationflags=flags,
            text=False,  # bytes 로 직접 다뤄 인코딩 일치 강제
        )
        return proc
    except Exception:
        return None


def _start_vw_drain(proc: subprocess.Popen) -> None:
    """helper 의 stdout(RC 응답) 을 백그라운드에서 소비. pipe buffer 가 가득 차
    helper 가 hang 하는 것을 방지. 메인 스레드는 발화 후 응답 대기 없이 즉시 반환."""
    global _VW_DRAIN_THREAD
    if _VW_DRAIN_THREAD is not None and _VW_DRAIN_THREAD.is_alive():
        return  # 이미 동작 중 (proc 재시작 시 자동으로 EOF 로 종료됨)

    def _drain():
        try:
            while True:
                line = proc.stdout.readline()
                if not line:  # EOF — proc 종료
                    break
        except Exception:
            pass

    th = threading.Thread(target=_drain, daemon=True)
    th.start()
    _VW_DRAIN_THREAD = th


def _speak_voicewith(text: str) -> bool:
    """보이스위드 helper 프록시로 발화 송신. 첫 호출에서 helper 가 자동 spawn.

    첫 발화는 RC 를 동기 검증해 보이스위드 동작 여부를 확인하고, 그 이후
    발화는 fire-and-forget — RC 응답 대기로 인한 메인 스레드 지연(매 발화마다
    10-50ms) 을 없애 TTS 반응 속도를 빠르게.
    """
    global _VW_PROXY, _VW_DISABLED, _VW_VERIFIED
    if _VW_DISABLED:
        return False
    with _VW_LOCK:
        # 살아 있는지 확인. 죽었으면 재시도 한 번.
        if _VW_PROXY is None or _VW_PROXY.poll() is not None:
            _VW_PROXY = _spawn_voicewith_proxy()
            _VW_VERIFIED = False
            if _VW_PROXY is None:
                _VW_DISABLED = True
                return False
        line = (text or "").replace("\r", " ").replace("\n", " ").strip()
        if not line:
            return False
        try:
            _VW_PROXY.stdin.write(line.encode("utf-8") + b"\n")
            _VW_PROXY.stdin.flush()
        except Exception:
            _VW_PROXY = None
            _VW_DISABLED = True
            return False

        # 첫 호출만 동기 RC 검증 — 이후엔 응답 대기 없이 즉시 반환.
        if _VW_VERIFIED:
            return True
        try:
            resp = _VW_PROXY.stdout.readline()
        except Exception:
            _VW_PROXY = None
            _VW_DISABLED = True
            return False
        try:
            text_resp = resp.decode("utf-8", errors="ignore").strip()
        except Exception:
            text_resp = ""
        if text_resp.startswith("RC:"):
            try:
                rc = int(text_resp.split(":", 1)[1])
            except Exception:
                rc = -1
            if rc == 0:
                _VW_VERIFIED = True
                _start_vw_drain(_VW_PROXY)
                return True
            _VW_DISABLED = True
            return False
        return False


def _cancel_voicewith() -> bool:
    if _VW_DISABLED or _VW_PROXY is None:
        return False
    with _VW_LOCK:
        if _VW_PROXY.poll() is not None:
            return False
        try:
            _VW_PROXY.stdin.write(b"__cancel__\n")
            _VW_PROXY.stdin.flush()
            # RC 응답은 백그라운드 드레인 스레드가 소비 — 동기 대기 안 함.
            return True
        except Exception:
            return False


# ─────────────── accessible_output2 (NVDA/보이스위드/JAWS/SAPI) ───────────────

# 라이브러리 자체가 nvdaControllerClient32/64.dll 을 동봉하고 있어 별도 번들·
# 경로 추적 없이 64-bit 환경에서도 NVDA/보이스위드와 통신 가능.
#
# 보이스위드 처리 — accessible_output2 의 `Auto()` 는 각 Output 의 is_active()
# 로 활성 스크린리더를 감지한다. NVDA Output 의 is_active() 는 nvdaController_
# testIfRunning() == 0 을 검사하는데, 보이스위드는 NVDA 와 RPC 엔드포인트가
# 일치하지 않아 이 검사가 실패하고 Auto() 가 SAPI 로 폴백한다 (사용자가 "보이스
# 위드 대신 SAPI 음성이 나온다"고 보고한 원인).
#
# 따라서: NVDA Output 인스턴스를 항상 만들어 두고 speakText 의 반환 코드(0=성공)
# 로 발화 성공 여부를 직접 판정한다. 보이스위드의 speakText 응답이 정상이면
# is_active() 결과와 상관없이 우리 발화가 들린다.
_AO2_NVDA = None
try:
    from accessible_output2.outputs.nvda import NVDA as _AO2NVDA  # type: ignore
    _AO2_NVDA = _AO2NVDA()
except Exception:
    _AO2_NVDA = None

_AO2_AUTO = None
try:
    from accessible_output2.outputs.auto import Auto as _AO2Auto  # type: ignore
    _AO2_AUTO = _AO2Auto()
except Exception:
    _AO2_AUTO = None


def _speak_ao2_nvda(text: str) -> bool:
    """accessible_output2 의 NVDA Output 을 통해 발화 시도.

    is_active() 결과를 신뢰하지 않고 lib.nvdaController_speakText 의 반환
    코드(0=성공) 로 직접 판정 — 보이스위드처럼 testIfRunning 응답이 없는
    NVDA 호환 스크린리더도 잡기 위함.
    """
    if _AO2_NVDA is None:
        return False
    try:
        lib = _AO2_NVDA.lib
        try:
            lib.nvdaController_cancelSpeech()
        except Exception:
            pass
        rc = lib.nvdaController_speakText(text)
        return rc == 0
    except Exception:
        return False


def _speak_ao2_auto(text: str) -> bool:
    """accessible_output2 의 Auto 를 통한 발화 — NVDA 우회 후 폴백."""
    if _AO2_AUTO is None:
        return False
    try:
        _AO2_AUTO.speak(text, interrupt=True)
        return True
    except Exception:
        return False


def _cancel_ao2() -> bool:
    cancelled = False
    if _AO2_NVDA is not None:
        try:
            _AO2_NVDA.lib.nvdaController_cancelSpeech()
            cancelled = True
        except Exception:
            pass
    if _AO2_AUTO is not None:
        try:
            try:
                _AO2_AUTO.silence()
            except AttributeError:
                _AO2_AUTO.speak("", interrupt=True)
            cancelled = True
        except Exception:
            pass
    return cancelled


# ───────────────────────── NVDA controllerClient ─────────────────────────

def _candidate_dll_paths() -> list[str]:
    """nvdaControllerClient64.dll / 32.dll 후보 경로를 우선순위 순으로."""
    names = ["nvdaControllerClient64.dll", "nvdaControllerClient32.dll"]
    dirs: list[str] = []

    # 1) PyInstaller 번들
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        dirs.append(os.path.join(meipass, "lib"))
        dirs.append(meipass)

    # 2) 실행 파일 디렉토리
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    for sub in ("", "lib", "bin"):
        d = os.path.join(exe_dir, sub) if sub else exe_dir
        if d not in dirs:
            dirs.append(d)

    # 3) NVDA / 보이스위드 설치 디렉토리 — 레지스트리 InstallLocation
    for inst in _find_nvda_compatible_install_locations():
        if inst not in dirs:
            dirs.append(inst)

    # 4) 시스템 PATH 의 일반 경로 + WindowsApps 류는 ctypes 가 알아서 검색
    candidates: list[str] = []
    for name in names:
        for d in dirs:
            p = os.path.join(d, name)
            if os.path.exists(p):
                candidates.append(p)
    # 마지막으로 이름만 — ctypes 가 시스템 PATH 에서 찾도록
    candidates.extend(names)
    return candidates


def _find_nvda_compatible_install_locations() -> list[str]:
    """레지스트리에서 NVDA 및 NVDA 기반 fork(보이스위드 등) 설치 경로 조회.

    NVDA 와 그 fork 들은 같은 `nvdaControllerClient*.dll` 을 설치 디렉토리
    또는 그 하위에 두는 경향이 있다. Uninstall 레지스트리 키의 InstallLocation
    을 모아 후보로 사용.
    """
    if platform.system() != "Windows":
        return []
    try:
        import winreg
    except Exception:
        return []

    # 후보 키 이름들. NVDA, VoiceWith, Voicewith, 한국어 표기 등 변형까지 포함.
    key_names = (
        "NVDA",
        "VoiceWith", "Voicewith", "voicewith",
        "보이스위드", "VoiceWithKor", "VoicewithKor",
    )
    parents = (
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    )

    found: list[str] = []

    # 1) 정확히 알려진 키 이름 직접 조회
    for hive, sub in parents:
        for name in key_names:
            try:
                with winreg.OpenKey(hive, f"{sub}\\{name}") as k:
                    val, _ = winreg.QueryValueEx(k, "InstallLocation")
                    if val and os.path.isdir(val) and val not in found:
                        found.append(val)
            except Exception:
                continue

    # 2) Uninstall 트리 전체를 훑어 DisplayName 으로 NVDA / Voice / 보이스위드
    # 패턴 매칭. 키 이름이 GUID 형태인 설치 사례도 잡기 위함.
    name_patterns = ("nvda", "voice", "보이스")
    for hive, sub in parents:
        try:
            root = winreg.OpenKey(hive, sub)
        except Exception:
            continue
        try:
            i = 0
            while True:
                try:
                    sub_name = winreg.EnumKey(root, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(root, sub_name) as k:
                        try:
                            display, _ = winreg.QueryValueEx(k, "DisplayName")
                        except FileNotFoundError:
                            continue
                        if not display:
                            continue
                        d_lower = display.lower()
                        if not any(p in d_lower for p in name_patterns):
                            continue
                        try:
                            install, _ = winreg.QueryValueEx(k, "InstallLocation")
                        except FileNotFoundError:
                            install = ""
                        if install and os.path.isdir(install) and install not in found:
                            found.append(install)
                except Exception:
                    continue
        finally:
            try:
                winreg.CloseKey(root)
            except Exception:
                pass

    return found


def _find_nvda_install_location() -> Optional[str]:
    """후방 호환용 — 첫 번째 후보 반환."""
    locs = _find_nvda_compatible_install_locations()
    return locs[0] if locs else None


def _load_nvda_dll():
    """nvdaControllerClient DLL 로드 시도. 성공 시 DLL 핸들 반환, 실패 시 None."""
    if platform.system() != "Windows":
        return None
    last_err: Optional[Exception] = None
    for path in _candidate_dll_paths():
        try:
            dll = ctypes.windll.LoadLibrary(path)
        except OSError as e:
            last_err = e
            continue
        try:
            # 함수 시그니처 — 안 정해 두면 wchar_p 인자가 잘려서 깨질 수 있음
            dll.nvdaController_speakText.argtypes = [ctypes.c_wchar_p]
            dll.nvdaController_speakText.restype = ctypes.c_int
            dll.nvdaController_cancelSpeech.argtypes = []
            dll.nvdaController_cancelSpeech.restype = ctypes.c_int
            dll.nvdaController_testIfRunning.argtypes = []
            dll.nvdaController_testIfRunning.restype = ctypes.c_int
            return dll
        except Exception as e:
            last_err = e
            continue
    if last_err is not None:
        # 디버그용으로만 보존 (조용한 폴백)
        global _NVDA_LOAD_ERROR
        _NVDA_LOAD_ERROR = str(last_err)
    return None


_NVDA_LOAD_ERROR: str = ""
_NVDA_DLL = _load_nvda_dll()


def _nvda_running() -> bool:
    """NVDA(또는 호환 스크린리더 — 보이스위드 등) 가 떠 있으면 True.

    `nvdaController_testIfRunning` 은 NVDA 가 등록한 특정 RPC 엔드포인트를
    검사한다. 보이스위드 같은 NVDA fork 는 이 함수가 제대로 응답하지 않는
    경우가 있어 detection 으로만 사용하고, 실제 발화는 speakText 결과로 판정.
    """
    if _NVDA_DLL is None:
        return False
    try:
        return _NVDA_DLL.nvdaController_testIfRunning() == 0
    except Exception:
        return False


def _speak_nvda(text: str) -> bool:
    """nvdaController_speakText 호출. testIfRunning 결과와 무관하게 반환값으로 판정.

    speakText 는 발화 성공 시 0 을 반환하고, 그렇지 않을 때 0 이 아닌 에러 코드.
    NVDA / 보이스위드 / 동등 fork 가 떠 있으면 0 을 받는다. 따라서 testIfRunning
    이 실패해도 (보이스위드처럼 RPC endpoint 이름이 다른 fork) speakText 가
    성공하면 그대로 사용한다.
    """
    if _NVDA_DLL is None:
        return False
    try:
        # cancelSpeech 는 실패해도 무시 — 발화 성공이 핵심.
        try:
            _NVDA_DLL.nvdaController_cancelSpeech()
        except Exception:
            pass
        result = _NVDA_DLL.nvdaController_speakText(text)
        # 0 = 성공. 그 외 코드(예: RPC_S_SERVER_UNAVAILABLE) 는 실패로 본다.
        return result == 0
    except Exception:
        return False


def _cancel_nvda() -> bool:
    if _NVDA_DLL is None:
        return False
    try:
        result = _NVDA_DLL.nvdaController_cancelSpeech()
        return result == 0
    except Exception:
        return False


# ───────────────────────── 센스리더 (COM 자동화) ─────────────────────────

_SENSE_APP = None


def _load_sense_reader():
    """센스리더 COM 객체를 한 번 시도해 캐싱. 실패 시 None."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client  # type: ignore
        return win32com.client.Dispatch("SenseReader.Application")
    except Exception:
        pass
    try:
        import comtypes.client  # type: ignore
        return comtypes.client.CreateObject("SenseReader.Application")
    except Exception:
        pass
    return None


_SENSE_APP = _load_sense_reader()


def _speak_sense_reader(text: str) -> bool:
    """센스리더 COM 발화. 모듈 import 시점에는 센스리더가 안 떠 있어 _SENSE_APP
    이 None 일 수 있다 — 첫 speak 호출 때 다시 시도해서 늦게 켜진 센스리더도 잡는다."""
    global _SENSE_APP
    if _SENSE_APP is None:
        _SENSE_APP = _load_sense_reader()
    if _SENSE_APP is None:
        return False
    try:
        try:
            _SENSE_APP.StopSpeaking()
        except Exception:
            pass
        _SENSE_APP.Speak(text)
        return True
    except Exception:
        # COM 객체가 stale 상태면 다음 호출에서 재로딩되도록 클리어
        _SENSE_APP = None
        return False


def _cancel_sense_reader() -> bool:
    if _SENSE_APP is None:
        return False
    try:
        _SENSE_APP.StopSpeaking()
        return True
    except Exception:
        return False


# ───────────────────────── Windows SAPI 5 (폴백) ─────────────────────────

_SAPI_VOICE = None


def _load_sapi():
    """Windows SAPI 5 음성 객체. 활성 스크린리더가 없을 때만 사용."""
    if platform.system() != "Windows":
        return None
    try:
        import comtypes.client  # type: ignore
        return comtypes.client.CreateObject("SAPI.SpVoice")
    except Exception:
        pass
    try:
        import win32com.client  # type: ignore
        return win32com.client.Dispatch("SAPI.SpVoice")
    except Exception:
        pass
    return None


def _speak_sapi(text: str) -> bool:
    """SAPI 음성 발화. 비동기(SVSFlagsAsync=1) + 이전 발화 중단(Purge=2) 동시 적용 = 3."""
    global _SAPI_VOICE
    if _SAPI_VOICE is None:
        _SAPI_VOICE = _load_sapi()
    if _SAPI_VOICE is None:
        return False
    try:
        _SAPI_VOICE.Speak(text, 3)  # 1 (Async) | 2 (PurgeBeforeSpeak)
        return True
    except Exception:
        return False


def _cancel_sapi() -> bool:
    global _SAPI_VOICE
    if _SAPI_VOICE is None:
        return False
    try:
        # 빈 문자열을 Purge 플래그로 발화하면 큐가 비워진다.
        _SAPI_VOICE.Speak("", 2)
        return True
    except Exception:
        return False


# ───────────────────────── 공개 API ─────────────────────────

def _estimate_speech_seconds(text: str) -> float:
    """발화 길이 추정 — 한국어 기준 음절 수에 비례.
    한글 1글자 ≈ 0.18s, 공백/숫자/영문은 짧게 잡는다. 최소 1.2s, 최대 5s."""
    if not text:
        return 0.0
    syllables = 0
    others = 0
    for ch in text:
        code = ord(ch)
        if 0xAC00 <= code <= 0xD7A3:  # 한글 음절
            syllables += 1
        elif ch.strip():
            others += 1
    secs = syllables * 0.18 + others * 0.06
    if secs < 1.2:
        secs = 1.2
    if secs > 5.0:
        secs = 5.0
    return secs


# 한 번이라도 "진짜" 스크린리더(보이스위드/NVDA/센스리더) 가 발화에 성공했는지
# 추적. True 면 이후 일시적 COM glitch 등으로 그 리더 호출이 False 를 반환해도
# SAPI/ao2_auto 로 폴백하지 않는다 — 사용자가 듣고 싶어하지 않는 SAPI 음성이
# 끼어들어 "센스리더 켜져 있는데도 가끔 SAPI 가 말함" 같은 증상 방지.
_REAL_READER_EVER_USED = False


def speak(text: str, wait: bool = False) -> bool:
    """스크린리더로 텍스트를 음성 출력.

    우선순위:
    1. 보이스위드 (32-bit DLL via PowerShell 프록시)
    2. accessible_output2 의 NVDA Output (NVDA·호환 fork)
    3. nvdaControllerClient 직접 호출
    4. 센스리더 COM 자동화
    5. accessible_output2 Auto() — JAWS/ZDSR 등
    6. SAPI 폴백

    1~4 중 한 번이라도 성공한 적이 있으면 그 이후엔 5/6 (SAPI) 으로 폴백하지
    않는다. 사용자가 진짜 스크린리더를 쓰고 있는데 일시적 실패로 SAPI 음성이
    끼어드는 것을 방지.

    이전 발화는 모두 중단하고 새 발화를 시작한다.

    wait=True 일 때: 발화 길이를 추정해 그만큼 sleep 한 뒤 반환.
        다음 호출자가 즉시 다른 speak() / 다이얼로그 변화 등으로 발화를
        잘라먹는 것을 막을 때 사용 (예: 인증 진입/완료 안내). UI 스레드를
        블록하므로 진행 중 비프음 같은 background 효과는 별도 스레드에서.
    """
    global _REAL_READER_EVER_USED
    if not text:
        return False
    text = str(text)
    spoken = False
    real_reader = False
    # 1) 보이스위드 (32-bit DLL via SysWOW64 PowerShell 프록시).
    if _speak_voicewith(text):
        spoken = True
        real_reader = True
    # 2) accessible_output2 의 NVDA Output 직접 호출 — NVDA·호환 fork.
    elif _speak_ao2_nvda(text):
        spoken = True
        real_reader = True
    # 3) 우리 자체 nvdaControllerClient 직접 호출 (DLL 다중 경로 폴백)
    elif _speak_nvda(text):
        spoken = True
        real_reader = True
    # 4) 센스리더 COM 자동화 — ao2_auto 보다 먼저. Auto() 는 활성 리더가
    #    없으면 SAPI 로 조용히 폴백하면서 True 를 반환해 버려, 사용자가
    #    듣고 싶어하는 센스리더로 가지 못한다.
    elif _speak_sense_reader(text):
        spoken = True
        real_reader = True
    # 진짜 스크린리더가 이전에 한 번이라도 동작했음 → SAPI 폴백 차단.
    # 일시적 COM glitch (특히 "인증 진행 중" 같은 백그라운드 스레드 발화에서
    # 가끔 발생) 가 SAPI 로 새는 것을 막는다.
    elif _REAL_READER_EVER_USED:
        return False
    # 5) accessible_output2 의 Auto() — 진짜 리더가 한 번도 안 잡힌 환경 한정
    elif _speak_ao2_auto(text):
        spoken = True
    # 6) SAPI 폴백 — 마지막
    elif _speak_sapi(text):
        spoken = True

    if real_reader:
        _REAL_READER_EVER_USED = True

    if spoken and wait:
        import time as _time
        _time.sleep(_estimate_speech_seconds(text))
    return spoken


def cancel_speech() -> bool:
    """현재 발화를 즉시 중단."""
    if platform.system() != "Windows":
        return False
    cancelled = False
    if _cancel_voicewith():
        cancelled = True
    if _cancel_ao2():
        cancelled = True
    if _cancel_nvda():
        cancelled = True
    if _cancel_sense_reader():
        cancelled = True
    if _cancel_sapi():
        cancelled = True
    return cancelled


def detect_active_reader() -> str:
    """현재 활성화된 스크린리더 이름을 반환. 진단·로그용.

    "AccessibleOutput2" / "NVDA" / "SenseReader" / "SAPI" / "none" 중 하나.
    """
    if _AO2_OUTPUT is not None:
        return "AccessibleOutput2"
    if _nvda_running():
        return "NVDA"
    if _SENSE_APP is not None:
        return "SenseReader"
    if _load_sapi() is not None:
        return "SAPI"
    return "none"
