@echo off
chcp 65001 >nul
SETLOCAL

REM 오늘 날짜 생성
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set TODAY=%%i

REM 내부 빌드용 이름(영문) + 실제 표시용 이름(한글)
set "INTERNAL_NAME=HalfetGetOrder_%TODAY%"
set "DISPLAY_NAME=하프전자 주문 수집기_%TODAY%"

echo TODAY        = %TODAY%
echo INTERNAL_NAME= %INTERNAL_NAME%
echo DISPLAY_NAME = %DISPLAY_NAME%
echo.

REM 현재 스크립트 위치
cd /d "%~dp0"

REM 가상환경 활성화
call venv\Scripts\activate.bat

REM 기존 빌드 파일 삭제
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

REM 아이콘 경로 설정
set "ICON_PATH=%~dp0icon\app.ico"

IF NOT EXIST "%ICON_PATH%" (
    echo [ERROR] 아이콘 파일을 찾을 수 없습니다:
    echo         "%ICON_PATH%"
    echo.
    pause
    goto :EOF
)

REM PyInstaller 빌드
pyinstaller --onefile --name "%INTERNAL_NAME%" --icon="%ICON_PATH%" entry.py

IF ERRORLEVEL 1 (
    echo.
    echo [ERROR] PyInstaller 빌드 중 오류가 발생했습니다.
    echo.
    goto CLEANUP
)


REM dist 안의 exe 파일을 한글 이름으로 변경
IF EXIST "dist\%INTERNAL_NAME%.exe" (
    ren "dist\%INTERNAL_NAME%.exe" "%DISPLAY_NAME%.exe"
)
    echo ✅ 빌드 완료! dist\%DISPLAY_NAME%.exe 를 확인하세요.
) ELSE (
    echo.
    echo [WARN] dist\%INTERNAL_NAME%.exe 파일을 찾을 수 없습니다.
    echo       PyInstaller 결과를 확인해 주세요.
)

REM JSON 파일 dist 폴더에 복사
copy /Y godo_add_goods_all.json dist\ >nul
copy /Y godo_goods_all.json dist\ >nul
copy /Y godo_base_specs.json dist\ >nul 2>nul

:CLEANUP
REM 8) 가상환경 비활성화 (안 돼도 그냥 넘어가게)
deactivate 2>nul

echo.
ENDLOCAL
pause
