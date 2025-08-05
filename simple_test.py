#!/usr/bin/env python3
# 간단한 OCR 및 테이블 추출 테스트

import os
import sys

def test_dependencies():
    """의존성 테스트"""
    print("=== 의존성 테스트 ===")
    
    # 1. Camelot 테스트
    try:
        import camelot
        print("✓ camelot-py 사용 가능")
        print(f"  버전: {camelot.__version__}")
    except ImportError as e:
        print(f"✗ camelot-py 실패: {e}")
        return False
    
    # 2. pdfplumber 테스트
    try:
        import pdfplumber
        print("✓ pdfplumber 사용 가능")
    except ImportError as e:
        print(f"✗ pdfplumber 실패: {e}")
        return False
    
    # 3. OCR 의존성 테스트
    import subprocess
    
    # tesseract 테스트
    try:
        result = subprocess.run(['tesseract', '--version'], 
                               capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            version_line = result.stdout.strip().split('\n')[0]
            print(f"✓ tesseract: {version_line}")
        else:
            print("✗ tesseract: 실행 실패")
            return False
    except FileNotFoundError:
        print("✗ tesseract: 설치되지 않음")
        return False
    except Exception as e:
        print(f"✗ tesseract: 오류 - {e}")
        return False
    
    # ocrmypdf 테스트
    try:
        result = subprocess.run(['ocrmypdf', '--version'], 
                               capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print(f"✓ ocrmypdf: {result.stdout.strip()}")
        else:
            print("⚠ ocrmypdf: 설치되어 있지만 실행 실패")
    except FileNotFoundError:
        print("✗ ocrmypdf: 설치되지 않음")
        return False
    except Exception as e:
        print(f"✗ ocrmypdf: 오류 - {e}")
        return False
    
    # ghostscript 테스트 (선택사항)
    try:
        result = subprocess.run(['gs', '--version'], 
                               capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print(f"✓ ghostscript: {result.stdout.strip()}")
        else:
            print("⚠ ghostscript: 설치되어 있지만 실행 실패")
    except FileNotFoundError:
        print("⚠ ghostscript: 설치되지 않음 (OCR 성능에 영향 있을 수 있음)")
    except Exception as e:
        print(f"⚠ ghostscript: 오류 - {e} (OCR 성능에 영향 있을 수 있음)")
    
    return True

def test_camelot():
    """Camelot 테스트"""
    print("\n=== Camelot 테스트 ===")
    
    import camelot
    
    # PDF 파일 찾기
    pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
    if not pdf_files:
        print("✗ 테스트할 PDF 파일이 없습니다.")
        return False
    
    test_pdf = pdf_files[0]
    print(f"테스트 PDF: {test_pdf}")
    
    try:
        print("1. 기본 방법으로 테이블 추출...")
        tables = camelot.read_pdf(test_pdf, pages='1')
        print(f"   결과: {len(tables)} 테이블 발견")
        
        if len(tables) > 0:
            print("   첫 번째 테이블 정보:")
            table = tables[0]
            print(f"     크기: {table.df.shape}")
            print(f"     정확도: {table.accuracy:.2f}")
            print("     내용 미리보기:")
            print(table.df.head(3))
        
        print("2. lattice 방법으로 테이블 추출...")
        tables_lattice = camelot.read_pdf(test_pdf, pages='1', flavor='lattice')
        print(f"   결과: {len(tables_lattice)} 테이블 발견")
        
        print("3. stream 방법으로 테이블 추출...")
        tables_stream = camelot.read_pdf(test_pdf, pages='1', flavor='stream')
        print(f"   결과: {len(tables_stream)} 테이블 발견")
        
        total_tables = len(tables) + len(tables_lattice) + len(tables_stream)
        if total_tables > 0:
            print(f"✓ 총 {total_tables}개 테이블 발견됨")
            return True
        else:
            print("⚠ 테이블을 찾을 수 없습니다. 이미지 기반 PDF일 가능성 높음")
            return False
        
    except Exception as e:
        print(f"✗ Camelot 테스트 실패: {e}")
        return False

def test_ocr():
    """OCR 테스트"""
    print("\n=== OCR 테스트 ===")
    
    # PDF 파일 찾기
    pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
    if not pdf_files:
        print("✗ 테스트할 PDF 파일이 없습니다.")
        return False
    
    test_pdf = pdf_files[0]
    print(f"테스트 PDF: {test_pdf}")
    
    import subprocess
    import tempfile
    
    try:
        # 임시 출력 파일 생성
        with tempfile.NamedTemporaryFile(suffix='_ocr.pdf', delete=False) as tmp_file:
            output_pdf = tmp_file.name
        
        print("OCR 처리 중...")
        cmd = [
            'ocrmypdf',
            '--language', 'eng+kor',
            test_pdf,
            output_pdf
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        
        if result.returncode == 0:
            print("✓ OCR 성공!")
            
            # OCR 후 파일 크기 확인
            original_size = os.path.getsize(test_pdf)
            ocr_size = os.path.getsize(output_pdf)
            print(f"  원본 크기: {original_size:,} bytes")
            print(f"  OCR 후 크기: {ocr_size:,} bytes")
            
            # OCR 후 Camelot 테스트
            print("OCR 후 테이블 추출 테스트...")
            import camelot
            tables_after_ocr = camelot.read_pdf(output_pdf, pages='1')
            print(f"  OCR 후 테이블 수: {len(tables_after_ocr)}")
            
            # 임시 파일 정리
            os.unlink(output_pdf)
            
            return True
        else:
            print(f"✗ OCR 실패: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ OCR 시간 초과")
        return False
    except Exception as e:
        print(f"✗ OCR 테스트 실패: {e}")
        return False

def main():
    """메인 함수"""
    print("=== PDF 처리 환경 테스트 ===")
    
    if not test_dependencies():
        print("\n✗ 의존성 테스트 실패 - 필요한 패키지가 설치되지 않았습니다.")
        return
    
    print("\n✓ 모든 의존성이 설치되었습니다!")
    
    # Camelot 테스트
    camelot_success = test_camelot()
    
    # OCR 테스트
    if not camelot_success:
        print("\n테이블을 찾을 수 없으므로 OCR 테스트를 진행합니다...")
        ocr_success = test_ocr()
        
        if ocr_success:
            print("\n✓ OCR 처리가 성공했습니다!")
        else:
            print("\n✗ OCR 처리도 실패했습니다.")
    else:
        print("\n✓ 테이블 추출이 성공했습니다!")
    
    print("\n=== 테스트 완료 ===")

if __name__ == '__main__':
    main()
