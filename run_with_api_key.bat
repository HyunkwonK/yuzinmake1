@echo off
echo ====================================================
echo PDF 번역기 실행 스크립트
echo ====================================================
echo.

REM DeepL API 키 설정
set DEEPL_API_KEY=b3125acc-3a44-4648-8b4d-5ca8e7350059:fx

echo ✓ DeepL API 키가 설정되었습니다.
echo.

REM Python 환경 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다.
    echo Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)

echo ✓ Python이 설치되어 있습니다.
echo.

REM test.py 파일 존재 확인
if not exist "test.py" (
    echo ❌ test.py 파일을 찾을 수 없습니다.
    echo 이 배치 파일을 test.py와 같은 폴더에 놓고 실행해주세요.
    pause
    exit /b 1
)

echo ✓ test.py 파일을 찾았습니다.
echo.

echo 🚀 PDF 번역기를 시작합니다...
echo 환경변수 DEEPL_API_KEY = %DEEPL_API_KEY%
echo.

REM Python 스크립트 실행
python test.py

echo.
echo 프로그램이 종료되었습니다.
pause
