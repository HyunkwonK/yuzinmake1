@echo off
chcp 65001 >nul
echo ====================================================
echo PDF 번역기 - 간단 실행
echo ====================================================
echo.

REM DeepL API 키 설정
set DEEPL_API_KEY=b3125acc-3a44-4648-8b4d-5ca8e7350059:fx

echo ✓ API 키 설정 완료
echo ✓ PDF 파일을 이 폴더에 넣고 실행하세요
echo.

REM Python 스크립트 실행
python test.py

pause
