@echo off
chcp 65001 > nul
echo ===================================
echo  초록멀티 빌드 환경 설정
echo  Python 3.12 가상환경 + 의존 패키지
echo ===================================

cd /d "%~dp0"

echo.
echo [1/3] Python 3.12 확인...
py -3.12 --version
if errorlevel 1 (
    echo.
    echo Python 3.12를 설치합니다...
    py install 3.12
    if errorlevel 1 (
        echo [오류] Python 3.12 설치 실패
        echo        https://www.python.org/downloads/release/python-3128/ 에서 직접 설치해 주세요.
        pause
        exit /b 1
    )
)

echo.
echo [2/3] 가상환경 생성...
if exist ".venv_py312" (
    echo 기존 가상환경이 있습니다. 유지합니다.
) else (
    py -3.12 -m venv .venv_py312
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패
        pause
        exit /b 1
    )
)

echo.
echo [3/3] 의존 패키지 설치...
".venv_py312\Scripts\python.exe" -m pip install --upgrade pip
".venv_py312\Scripts\python.exe" -m pip install wxPython==4.2.2 requests beautifulsoup4 lxml cryptography pywin32 comtypes pyinstaller
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo.
echo ===================================
echo  환경 설정 완료!
echo  이제 build.bat 을 실행해서 빌드하세요.
echo ===================================
pause
