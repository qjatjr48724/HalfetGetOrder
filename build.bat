@echo off
chcp 65001 >nul
SETLOCAL

REM ─────────────────────────────────────────
REM 메인 메뉴
REM ─────────────────────────────────────────
:MENU
cls
echo ========================================
echo   하프전자 주문수집기 - 빌드 / 설정 도구
echo ========================================
echo   사용 안내:
echo     1번 - 새 실행파일(exe)을 만들 때 사용
echo     2번 - 쿠팡 / 고도몰 API 키가 변경됐을 때 사용
echo     3번 - 방금(또는 최근) 빌드된 exe를 실행할 때 사용
echo     4번 - 창을 닫고 종료할 때 사용
echo ========================================
echo   1. 프로그램 빌드
echo   2. API 키 변경 (keys.py + .env)
echo   3. 최근 빌드된 exe 실행
echo   4. 아무 것도 안 하고 종료
echo ========================================
set "MENU_CHOICE="
set /p MENU_CHOICE=번호를 선택하세요 (1-4^) : 

if "%MENU_CHOICE%"=="1" goto BUILD
if "%MENU_CHOICE%"=="2" goto CHANGE_KEYS
if "%MENU_CHOICE%"=="3" goto RUN_LAST
if "%MENU_CHOICE%"=="4" goto END

echo.
echo [WARN] 잘못된 입력입니다. 1~4 중에서 선택해주세요.
pause
goto MENU


REM ─────────────────────────────────────────
REM 2. API 키 변경
REM ─────────────────────────────────────────
:CHANGE_KEYS
cls
echo [INFO] API 키 변경을 시작합니다.
echo.

REM 스크립트 위치로 이동
cd /d "%~dp0"

REM 가상환경 활성화 (venv 가정)
call venv\Scripts\activate.bat

python update_keys.py

REM 가상환경 비활성화 (실패해도 무시)
deactivate 2>nul

echo.
echo [INFO] API 키 변경 작업이 완료되었습니다.
echo.
pause
goto MENU


REM ─────────────────────────────────────────
REM 1. 빌드
REM ─────────────────────────────────────────
:BUILD
cls
echo [INFO] PyInstaller 빌드를 시작합니다.
echo.

REM 스크립트 위치
cd /d "%~dp0"

REM 오늘 날짜 생성 (파일명에 사용)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set TODAY=%%i

REM 내부 빌드용 이름(영문) + 실제 표시용 이름(한글)
set "INTERNAL_NAME=HalfetGetOrder_%TODAY%"
set "DISPLAY_NAME=하프전자 주문 수집기_%TODAY%"

echo TODAY        = %TODAY%
echo INTERNAL_NAME= %INTERNAL_NAME%
echo DISPLAY_NAME = %DISPLAY_NAME%
echo.

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
    goto MENU
)

REM PyInstaller 빌드
pyinstaller --onefile --name "%INTERNAL_NAME%" --icon="%ICON_PATH%" entry.py

IF ERRORLEVEL 1 (
    echo.
    echo [ERROR] PyInstaller 빌드 중 오류가 발생했습니다.
    echo.
    goto MENU
)

REM dist 안의 exe 파일을 한글 이름으로 변경
IF EXIST "dist\%INTERNAL_NAME%.exe" (
    ren "dist\%INTERNAL_NAME%.exe" "%DISPLAY_NAME%.exe"
    echo.
    echo ✅ 빌드 완료! dist\%DISPLAY_NAME%.exe 를 확인하세요.
) ELSE (
    echo.
    echo [WARN] dist\%INTERNAL_NAME%.exe 파일을 찾을 수 없습니다.
    echo       PyInstaller 결과를 확인해 주세요.
)

echo.
pause
goto MENU


REM ─────────────────────────────────────────
REM 3. 최근 빌드된 exe 실행
REM ─────────────────────────────────────────
:RUN_LAST
cls
echo [INFO] 최근 빌드된 exe 를 찾는 중입니다...
echo.

cd /d "%~dp0"

IF NOT EXIST "dist" (
    echo [WARN] dist 폴더가 존재하지 않습니다.
    echo        먼저 1번으로 프로그램을 빌드해주세요.
    echo.
    pause
    goto MENU
)

set "LATEST_EXE="

for /f "delims=" %%F in ('dir /b /o-d "dist\*.exe" 2^>nul') do (
    if not defined LATEST_EXE set "LATEST_EXE=%%F"
)

IF NOT DEFINED LATEST_EXE (
    echo [WARN] dist 폴더에 exe 파일이 없습니다.
    echo        먼저 1번으로 프로그램을 빌드해주세요.
    echo.
    pause
    goto MENU
)

echo [INFO] 최근 빌드된 파일: dist\%LATEST_EXE%
echo.
set /p RUN_NOW=지금 이 파일을 실행할까요? (Y/N, 엔터=예) : 

if /i "%RUN_NOW%"=="n" (
    echo.
    echo 실행하지 않고 메뉴로 돌아갑니다.
    echo.
    pause
    goto MENU
)

echo.
echo [INFO] 프로그램을 실행합니다...
start "" "dist\%LATEST_EXE%"
echo.
echo 실행 요청을 보냈습니다. (이 창을 닫아도 프로그램은 별도 실행됩니다.)
echo.
pause
goto MENU


REM ─────────────────────────────────────────
REM 종료
REM ─────────────────────────────────────────
:END
ENDLOCAL
exit /b
