@echo off
echo ========================================
echo       PDF 처리기 (Docker 버전)
echo ========================================
echo.

REM 현재 디렉토리 확인
echo 현재 작업 디렉토리: %CD%
echo.

REM PDF 파일 확인
echo PDF 파일 검색 중...
dir *.pdf /b >nul 2>&1
if errorlevel 1 (
    echo 오류: 현재 폴더에 PDF 파일이 없습니다.
    echo PDF 파일을 이 폴더에 넣고 다시 실행해주세요.
    pause
    exit /b 1
)

echo 발견된 PDF 파일들:
dir *.pdf /b
echo.

REM Docker 확인
echo Docker 확인 중...
docker --version >nul 2>&1
if errorlevel 1 (
    echo 오류: Docker가 설치되지 않았거나 실행되지 않습니다.
    echo Docker Desktop을 설치하고 실행한 후 다시 시도해주세요.
    pause
    exit /b 1
)
echo Docker 사용 가능
echo.

REM Docker 이미지 확인
echo PDF 처리기 이미지 확인 중...
docker images pdf-processor --format "table {{.Repository}}\t{{.Tag}}\t{{.Size}}" | find "pdf-processor" >nul
if errorlevel 1 (
    echo 오류: pdf-processor Docker 이미지를 찾을 수 없습니다.
    echo 먼저 다음 명령으로 이미지를 빌드해주세요:
    echo docker build -t pdf-processor .
    pause
    exit /b 1
)
echo PDF 처리기 이미지 준비됨
echo.

REM 실행
echo PDF 처리 시작...
echo ========================================
docker run --rm -v "%CD%":/app/input -v "%CD%":/app/output pdf-processor

echo.
echo ========================================
echo 처리 완료! 결과 파일을 확인해주세요.
echo ========================================
pause
