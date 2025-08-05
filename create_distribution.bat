@echo off
echo ========================================
echo    Docker 이미지 배포 패키지 생성기
echo ========================================
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

echo Docker 이미지를 tar 파일로 저장 중...
docker save pdf-processor:latest -o pdf-processor-image.tar

if exist pdf-processor-image.tar (
    echo.
    echo ========================================
    echo 성공! 배포 패키지가 생성되었습니다.
    echo.
    echo 파일: pdf-processor-image.tar
    echo 크기: 
    for %%I in (pdf-processor-image.tar) do echo %%~zI bytes
    echo.
    echo 다른 컴퓨터에서 사용하려면:
    echo 1. Docker Desktop 설치
    echo 2. docker load -i pdf-processor-image.tar
    echo 3. run_pdf_processor.bat 실행
    echo ========================================
) else (
    echo 오류: 이미지 저장에 실패했습니다.
)

pause
