#!/usr/bin/env python3
"""
PDF 테이블 재처리 스크립트
기존에 처리된 PDF들의 테이블을 다시 추출하여 덮어씁니다.
"""

import os
import shutil
import pandas as pd
from datetime import datetime

def reprocess_pdf_tables():
    """기존 PDF들의 테이블을 다시 처리"""
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    origin_dir = os.path.join(base_dir, 'Medical', 'Context', 'Origin')
    temperal_pdf_dir = os.path.join(base_dir, 'temperal_pdf')
    excel_file = os.path.join(base_dir, 'Medical_Table_Results.xlsx')
    
    print("🔄 PDF 테이블 재처리 시작")
    print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # temperal_pdf 디렉토리 생성
    os.makedirs(temperal_pdf_dir, exist_ok=True)
    
    # 기존 엑셀 파일에서 PDF 파일 정보 읽기
    try:
        main_df = pd.read_excel(excel_file, sheet_name='Main Results')
        pdf_entries = main_df[main_df['URL'].str.startswith('PDF_FILE:')]
        print(f"📋 엑셀 파일에서 {len(pdf_entries)}개의 PDF 항목 발견")
    except Exception as e:
        print(f"❌ 엑셀 파일 읽기 실패: {e}")
        return False
    
    if len(pdf_entries) == 0:
        print("⚠️  처리할 PDF가 없습니다.")
        return True
    
    # Origin 디렉토리에서 PDF 파일 찾아서 temperal_pdf로 복사
    pdf_files_copied = 0
    
    for idx, row in pdf_entries.iterrows():
        origin_number = row['Origin Number']
        pdf_file_path = os.path.join(origin_dir, f'M_origin_{origin_number}.pdf')
        
        if os.path.exists(pdf_file_path):
            # PDF 파일명 추출
            url_field = row['URL']
            pdf_filename = url_field.replace('PDF_FILE: ', '').strip()
            
            # temperal_pdf로 복사
            target_path = os.path.join(temperal_pdf_dir, pdf_filename)
            shutil.copy2(pdf_file_path, target_path)
            
            print(f"📄 복사됨: {pdf_filename} (Origin {origin_number})")
            pdf_files_copied += 1
        else:
            print(f"❌ 파일 없음: {pdf_file_path}")
    
    print(f"\n📊 총 {pdf_files_copied}개의 PDF 파일을 temperal_pdf로 복사했습니다.")
    
    if pdf_files_copied > 0:
        print("\n🚀 이제 pdf_processor.py를 실행하여 테이블을 다시 추출합니다...")
        print("💡 기존 테이블 이미지들이 덮어씌워집니다.")
        print("\n실행 명령어:")
        print("python pdf_processor.py")
    
    return True

if __name__ == "__main__":
    try:
        success = reprocess_pdf_tables()
        if success:
            print(f"\n✅ 재처리 준비 완료!")
        else:
            print(f"\n❌ 재처리 준비 실패!")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")