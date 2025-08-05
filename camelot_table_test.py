#!/usr/bin/env python3
# 실제 PDF에서 camelot 테이블 추출 테스트

import camelot
import os

print("=== Camelot 테이블 추출 테스트 ===")

# PDF 파일 찾기
pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
print(f"발견된 PDF 파일들: {pdf_files}")

if not pdf_files:
    print("테스트할 PDF 파일이 없습니다.")
    exit(1)

# 첫 번째 PDF 파일로 테스트
test_pdf = pdf_files[0]
print(f"\n테스트 PDF: {test_pdf}")

try:
    # 다양한 방법으로 테이블 추출 시도
    print("\n1. 기본 방법으로 테이블 추출...")
    tables1 = camelot.read_pdf(test_pdf, pages='all')
    print(f"   결과: {len(tables1)} 테이블 발견")
    
    print("\n2. lattice 방법으로 테이블 추출...")
    tables2 = camelot.read_pdf(test_pdf, pages='all', flavor='lattice')
    print(f"   결과: {len(tables2)} 테이블 발견")
    
    print("\n3. stream 방법으로 테이블 추출...")
    tables3 = camelot.read_pdf(test_pdf, pages='all', flavor='stream')
    print(f"   결과: {len(tables3)} 테이블 발견")
    
    # 가장 많은 테이블을 찾은 방법 선택
    best_tables = tables1
    best_method = "기본"
    
    if len(tables2) > len(best_tables):
        best_tables = tables2
        best_method = "lattice"
    
    if len(tables3) > len(best_tables):
        best_tables = tables3
        best_method = "stream"
    
    print(f"\n가장 좋은 결과: {best_method} 방법으로 {len(best_tables)} 테이블 발견")
    
    # 테이블 정보 출력
    if len(best_tables) > 0:
        print("\n=== 발견된 테이블 정보 ===")
        for i, table in enumerate(best_tables):
            print(f"테이블 {i+1}:")
            print(f"  페이지: {table.page}")
            print(f"  크기: {table.df.shape}")
            print(f"  정확도: {table.accuracy:.2f}")
            print(f"  화이트스페이스: {table.whitespace:.2f}")
            print(f"  내용 미리보기:")
            print(table.df.head(3))
            print("-" * 50)
    else:
        print("\n⚠ 테이블을 찾을 수 없습니다.")
        print("가능한 원인:")
        print("1. PDF가 이미지 기반일 수 있습니다 (OCR 필요)")
        print("2. 테이블이 표준 형식이 아닐 수 있습니다")
        print("3. 파라미터 조정이 필요할 수 있습니다")

except Exception as e:
    print(f"✗ camelot 테이블 추출 실패: {e}")
    import traceback
    traceback.print_exc()

print("\n=== 테스트 완료 ===")
