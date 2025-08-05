#!/usr/bin/env python3
# camelot 문제 진단 스크립트

print("=== Camelot 문제 진단 ===")

# 1. 기본 import 테스트
try:
    import camelot
    print("✓ camelot 기본 import 성공")
    print(f"  버전: {camelot.__version__}")
    print(f"  설치 위치: {camelot.__file__}")
except Exception as e:
    print(f"✗ camelot 기본 import 실패: {e}")
    exit(1)

# 2. core 모듈 확인
try:
    from camelot import core
    print("✓ camelot.core import 성공")
except Exception as e:
    print(f"✗ camelot.core import 실패: {e}")

# 3. TableList 확인 - 여기가 문제
try:
    from camelot.core import TableList
    print("✓ TableList import 성공")
except Exception as e:
    print(f"✗ TableList import 실패: {e}")
    print("  이것이 주요 문제입니다!")

# 4. 다른 방법으로 TableList 찾기
try:
    import camelot.core
    print("\ncamelot.core 모듈 내용:")
    print(dir(camelot.core))
except Exception as e:
    print(f"core 모듈 내용 확인 실패: {e}")

# 5. 의존성 확인
print("\n=== 의존성 확인 ===")
dependencies = ['pandas', 'numpy', 'cv2', 'pdfplumber']

for dep in dependencies:
    try:
        if dep == 'cv2':
            import cv2
            print(f"✓ opencv-python: {cv2.__version__}")
        else:
            module = __import__(dep)
            version = getattr(module, '__version__', 'unknown')
            print(f"✓ {dep}: {version}")
    except ImportError:
        print(f"✗ {dep}: 설치되지 않음")

# 6. camelot 직접 사용 테스트
print("\n=== camelot 직접 사용 테스트 ===")
try:
    # PDF 파일이 있는지 확인
    import os
    pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
    if pdf_files:
        test_pdf = pdf_files[0]
        print(f"테스트 PDF: {test_pdf}")
        
        # read_pdf 함수 직접 호출
        tables = camelot.read_pdf(test_pdf, pages='1')
        print(f"✓ camelot.read_pdf 성공: {len(tables)} 테이블 발견")
    else:
        print("테스트할 PDF 파일이 없습니다.")
        
except Exception as e:
    print(f"✗ camelot 사용 테스트 실패: {e}")
    import traceback
    traceback.print_exc()

print("\n=== 진단 완료 ===")
