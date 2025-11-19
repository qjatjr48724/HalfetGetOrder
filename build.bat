@echo off
SETLOCAL

REM 현재 스크립트 위치로 이동
cd /d "%~dp0"

REM 가상환경 활성화
call venv\Scripts\activate.bat

REM 기존 빌드 파일 삭제
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q HalfetGetOrder.spec 2>nul

REM PyInstaller 빌드
pyinstaller ^
  --onefile ^
  --name HalfetGetOrder ^
  entry.py

REM 가상환경 비활성화
deactivate

echo.
echo 빌드 완료! dist\HalfetGetOrder.exe 를 확인하세요.
echo.

ENDLOCAL
pause
