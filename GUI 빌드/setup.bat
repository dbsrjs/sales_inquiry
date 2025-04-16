@echo off
SETLOCAL

REM -----------------------------
REM 1. Python 설치 여부 확인
REM -----------------------------
where python >nul 2>nul
IF ERRORLEVEL 1 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo https://www.python.org/downloads/ 에서 Python 3.11 이상을 설치해주세요.
    pause
    exit /b
)

REM -----------------------------
REM 2. pip 설치 확인
REM -----------------------------
python -m ensurepip --default-pip
python -m pip install --upgrade pip

REM -----------------------------
REM 3. 필요한 패키지 설치
REM -----------------------------
echo ✅ 필요한 패키지를 설치 중입니다...
python -m pip install pandas openpyxl xlrd pyinstaller

REM tkinter는 기본 내장이라 별도 설치 필요 없음

REM -----------------------------
REM 4. 스크립트 빌드
REM -----------------------------
echo ✅ PyInstaller로 실행 파일 생성 중입니다...
pyinstaller --noconfirm --onefile --windowed make_excel_GUI.py

echo ✅ 설치 완료! dist\main.exe를 실행하세요.
pause
