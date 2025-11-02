#!/usr/bin/env python3
"""
PDF 테이블 강제 재처리 스크립트
기존 PDF들의 테이블을 강제로 다시 추출하여 덮어씁니다.
"""

import os
import sys
import shutil
import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime

class PDFTableReprocessor:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.origin_dir = os.path.join(self.base_dir, 'Medical', 'Context', 'Origin')
        self.table_dir = os.path.join(self.base_dir, 'Medical', 'Table')
        self.excel_file = os.path.join(self.base_dir, 'Medical_Table_Results.xlsx')
        
        # 디렉토리 생성
        os.makedirs(self.table_dir, exist_ok=True)

    def extract_tables_from_pdf(self, pdf_path, origin_number):
        """PDF에서 테이블 추출 (수정된 버전)"""
        try:
            print(f"PDF에서 테이블 추출 시작: {pdf_path}")
            
            pdf_document = fitz.open(pdf_path)
            table_info = []
            
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                
                # 테이블 검색
                try:
                    tables = page.find_tables()
                    table_list = list(tables) if tables else []
                except Exception as table_find_error:
                    print(f"페이지 {page_num + 1}에서 테이블 검색 실패: {table_find_error}")
                    table_list = []
                
                if table_list:
                    print(f"페이지 {page_num + 1}에서 {len(table_list)}개의 테이블을 발견했습니다.")
                    
                    for table_idx, table in enumerate(table_list):
                        try:
                            # 테이블 영역 추출
                            table_rect = table.bbox
                            
                            # bbox가 tuple인 경우 Rect 객체로 변환
                            if isinstance(table_rect, tuple):
                                table_rect = fitz.Rect(table_rect)
                            
                            # 테이블 영역 확장 (패딩 추가하여 잘림 방지)
                            padding = 20  # 20 포인트 패딩
                            expanded_rect = fitz.Rect(
                                max(0, table_rect.x0 - padding),  # 왼쪽 패딩
                                max(0, table_rect.y0 - padding),  # 위쪽 패딩
                                min(page.rect.x1, table_rect.x1 + padding),  # 오른쪽 패딩 (페이지 경계 제한)
                                min(page.rect.y1, table_rect.y1 + padding)   # 아래쪽 패딩 (페이지 경계 제한)
                            )
                            
                            # 테이블 영역을 이미지로 캡처 (더 높은 해상도)
                            matrix = fitz.Matrix(400/72, 400/72)  # 400 DPI로 증가
                            pix = page.get_pixmap(matrix=matrix, clip=expanded_rect)
                            
                            # 테이블 이미지 저장 (기존 파일 덮어쓰기)
                            table_filename = f"M_table_{origin_number}_{table_idx}.png"
                            table_path = os.path.join(self.table_dir, table_filename)
                            pix.save(table_path)
                            
                            # 테이블 데이터 추출
                            try:
                                table_data = table.extract()
                            except Exception as extract_error:
                                print(f"테이블 데이터 추출 실패, 기본값 사용: {extract_error}")
                                table_data = []
                            
                            # 미리보기 텍스트 생성
                            preview_text = ""
                            if table_data:
                                # 첫 2-3 행의 텍스트를 미리보기로 사용
                                for row_idx, row in enumerate(table_data[:3]):
                                    if row:
                                        row_text = " | ".join([str(cell) if cell else "" for cell in row])
                                        preview_text += row_text + " "
                                        if len(preview_text) > 150:
                                            break
                            
                            if len(preview_text) > 200:
                                preview_text = preview_text[:200] + "..."
                            elif not preview_text:
                                preview_text = f"Page {page_num + 1} Table {table_idx + 1}"
                            
                            table_info.append({
                                'table_number': len(table_info),
                                'filename': table_path,
                                'page_number': page_num + 1,
                                'table_index_in_page': table_idx,
                                'preview_text': preview_text.strip(),
                                'rows': len(table_data) if table_data else 0,
                                'columns': len(table_data[0]) if table_data and len(table_data) > 0 else 0,
                                'size': f"{len(table_data) if table_data else 0}x{len(table_data[0]) if table_data and len(table_data) > 0 else 0}",
                                'image_size': f"{int((expanded_rect.x1 - expanded_rect.x0) * 400/72)}x{int((expanded_rect.y1 - expanded_rect.y0) * 400/72)}",
                                'position': f"Page {page_num + 1}"
                            })
                            
                            print(f"✅ 테이블 재추출 완료: {table_filename} (페이지 {page_num + 1})")
                            
                        except Exception as table_error:
                            print(f"❌ 페이지 {page_num + 1}의 테이블 {table_idx} 추출 실패: {table_error}")
                            continue
            
            pdf_document.close()
            print(f"총 {len(table_info)}개의 테이블을 재추출했습니다.")
            return table_info
            
        except Exception as e:
            print(f"PDF 테이블 추출 실패: {e}")
            return []

    def reprocess_all_pdfs(self):
        """모든 PDF의 테이블을 재처리"""
        print("🔄 PDF 테이블 강제 재처리 시작")
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)
        
        # 기존 엑셀 파일에서 PDF 정보 읽기
        try:
            main_df = pd.read_excel(self.excel_file, sheet_name='Main Results')
            pdf_entries = main_df[main_df['URL'].str.startswith('PDF_FILE:')]
            print(f"📋 처리할 PDF: {len(pdf_entries)}개")
        except Exception as e:
            print(f"❌ 엑셀 파일 읽기 실패: {e}")
            return False
        
        total_tables_reprocessed = 0
        
        for idx, (_, row) in enumerate(pdf_entries.iterrows(), 1):
            origin_number = row['Origin Number']
            pdf_filename = row['URL'].replace('PDF_FILE: ', '').strip()
            pdf_path = os.path.join(self.origin_dir, f'M_origin_{origin_number}.pdf')
            
            print(f"\n진행상황: {idx}/{len(pdf_entries)}")
            print(f"{'='*50}")
            print(f"재처리 중: {pdf_filename}")
            print(f"Origin Number: {origin_number}")
            print(f"{'='*50}")
            
            if os.path.exists(pdf_path):
                # 테이블 재추출
                table_info = self.extract_tables_from_pdf(pdf_path, origin_number)
                total_tables_reprocessed += len(table_info)
                
                print(f"✅ PDF 재처리 완료: {len(table_info)}개 테이블 재추출")
            else:
                print(f"❌ PDF 파일 없음: {pdf_path}")
        
        print(f"\n{'='*70}")
        print(f"🎉 모든 PDF 재처리 완료!")
        print(f"📊 총 {total_tables_reprocessed}개의 테이블이 재추출되었습니다.")
        print(f"💡 기존 테이블 이미지들이 더 나은 품질로 덮어씌워졌습니다.")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*70}")
        
        return True

def main():
    """메인 실행 함수"""
    try:
        reprocessor = PDFTableReprocessor()
        
        print("⚠️  이 작업은 기존의 모든 테이블 이미지를 새로운 버전으로 덮어씁니다.")
        print("💡 오른쪽 잘림 문제가 해결된 더 나은 품질의 테이블 이미지가 생성됩니다.")
        print("\n계속하시겠습니까? (y/N): ", end="")
        
        try:
            user_input = input().strip().lower()
            if user_input not in ['y', 'yes', '예', 'ㅇ']:
                print("재처리가 취소되었습니다.")
                return False
        except KeyboardInterrupt:
            print("\n재처리가 취소되었습니다.")
            return False
        
        success = reprocessor.reprocess_all_pdfs()
        return success
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)