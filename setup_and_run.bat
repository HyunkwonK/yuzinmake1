@echo off
chcp 65001 >nul
echo ====================================================
echo PDF 번역기 - 자동 설치 및 실행 스크립트
echo ====================================================
echo.

REM 관리자 권한 확인
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠ 일부 기능을 위해 관리자 권한이 필요할 수 있습니다.
    echo 만약 설치 중 오류가 발생하면 '관리자 권한으로 실행'을 시도해주세요.
    echo.
)

REM DeepL API 키 설정
set DEEPL_API_KEY=b3125acc-3a44-4648-8b4d-5ca8e7350059:fx
echo ✓ DeepL API 키가 설정되었습니다.
echo.

REM Python 설치 확인
echo [1/5] Python 환경 확인 중...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다.
    echo.
    echo Python 설치 방법:
    echo 1. https://www.python.org/downloads/ 에서 Python 다운로드
    echo 2. 설치 시 'Add Python to PATH' 체크박스 선택 필수
    echo 3. 설치 후 컴퓨터 재시작
    echo.
    pause
    exit /b 1
)
echo ✓ Python이 설치되어 있습니다.

REM pip 업그레이드
echo [2/5] pip 업그레이드 중...
python -m pip install --upgrade pip --quiet

REM 필수 Python 패키지 설치
echo [3/5] 필수 Python 패키지 설치 중...
echo 이 과정은 몇 분 소요될 수 있습니다...
python -m pip install pandas openpyxl xlsxwriter requests pdfplumber camelot-py[cv] --quiet
if %errorlevel% neq 0 (
    echo ⚠ 일부 패키지 설치에 실패했습니다. 기본 기능은 동작할 수 있습니다.
)

REM OCR 패키지 설치 (선택사항)
echo [4/5] OCR 패키지 설치 중...
python -m pip install ocrmypdf --quiet
if %errorlevel% neq 0 (
    echo ⚠ OCR 패키지 설치에 실패했습니다. 텍스트 추출만 가능합니다.
)

REM Chocolatey 확인 및 OCR 도구 설치 안내
echo [5/5] OCR 도구 확인 중...
choco --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠ Chocolatey가 설치되지 않았습니다.
    echo.
    echo OCR 기능을 위한 추가 설치 (선택사항):
    echo 1. 관리자 권한으로 PowerShell 실행
    echo 2. 다음 명령어 실행:
    echo    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    echo 3. 설치 후: choco install tesseract ghostscript -y
    echo.
) else (
    echo ✓ Chocolatey가 설치되어 있습니다.
    echo OCR 도구 설치를 시도합니다...
    choco install tesseract ghostscript -y >nul 2>&1
    if %errorlevel% equ 0 (
        echo ✓ OCR 도구가 설치되었습니다.
    ) else (
        echo ⚠ OCR 도구 설치에 실패했지만 계속 진행합니다.
    )
)

echo.
echo ====================================================
echo 설치가 완료되었습니다!
echo ====================================================
echo.

REM test.py 파일 존재 확인
if not exist "test.py" (
    echo ❌ test.py 파일을 찾을 수 없습니다.
    echo 이 배치 파일을 test.py와 같은 폴더에 놓고 실행해주세요.
    pause
    exit /b 1
)

echo 사용 방법:
echo 1. PDF 파일을 이 폴더에 복사하세요
echo 2. 아래 버튼을 누르면 자동으로 처리됩니다
echo 3. 완료되면 '_final.xlsx' 파일이 생성됩니다
echo.

set /p answer="PDF 번역기를 시작하시겠습니까? (Y/N): "
if /i "%answer%"=="Y" (
    echo.
    echo 🚀 PDF 번역기를 시작합니다...
    echo 환경변수 DEEPL_API_KEY = %DEEPL_API_KEY%
    echo.
    
    REM Python 스크립트 실행
    python test.py
    
    echo.
    echo 작업이 완료되었습니다!
) else (
    echo 나중에 'run_with_api_key.bat' 파일을 실행하세요.
)

echo.
pause
