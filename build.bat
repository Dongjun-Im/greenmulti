@echo off
chcp 65001 > nul
echo ===================================
echo  초록멀티 빌드 (Python 3.12 기반)
echo  - Win10/Win11 모두 호환
echo ===================================

cd /d "%~dp0"

if not exist ".venv_py312\Scripts\python.exe" (
    echo [오류] Python 3.12 가상환경이 없습니다.
    echo        먼저 setup.bat을 실행해 주세요.
    pause
    exit /b 1
)

echo.
echo [1/2] 이전 빌드 정리...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo [2/2] PyInstaller 실행...
".venv_py312\Scripts\python.exe" -m PyInstaller chorok_multi.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo.
echo ===================================
echo  빌드 완료: dist\초록멀티\
echo ===================================
pause
