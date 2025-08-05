#!/usr/bin/env python3
# Camelot으로 찾은 테이블 상세 확인

import camelot
import pandas as pd

print("=== Camelot 테이블 추출 결과 ===")

# PDF에서 테이블 추출
pdf_file = "iti Canned Ingredient Dec Cat & Dog Food August 2024.pdf"
tables = camelot.read_pdf(pdf_file, pages='all')

print(f"총 {len(tables)}개의 테이블을 찾았습니다.")

# 각 테이블 정보 출력
for i, table in enumerate(tables):
    print(f"\n--- 테이블 {i+1} ---")
    print(f"페이지: {table.page}")
    print(f"크기: {table.df.shape[0]} 행 x {table.df.shape[1]} 열")
    print(f"정확도: {table.accuracy:.2f}")
    print(f"화이트스페이스: {table.whitespace:.2f}")
    
    # 테이블 내용 미리보기
    print("내용 미리보기:")
    print(table.df.head(5))
    print("-" * 50)

# Excel로 저장
if len(tables) > 0:
    print("\nExcel 파일로 저장 중...")
    
    with pd.ExcelWriter("camelot_extracted_tables.xlsx", engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            sheet_name = f"Table_{i+1}_Page_{table.page}"
            table.df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"테이블 {i+1} → {sheet_name} 시트에 저장")
    
    print("✓ camelot_extracted_tables.xlsx 파일이 생성되었습니다!")

print("\n=== 완료 ===")
