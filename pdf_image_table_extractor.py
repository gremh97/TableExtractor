#!/usr/bin/env python3
"""
PDF to PNG 변환 후 이미지 기반 테이블 추출
PDF를 PNG로 변환하여 이미지에서 테이블 영역을 감지하고 추출합니다.
"""

import os
import cv2
import numpy as np
import fitz  # PyMuPDF
from PIL import Image
import pandas as pd
from datetime import datetime

class PDFImageTableExtractor:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.origin_dir = os.path.join(self.base_dir, 'Medical', 'Context', 'Origin')
        self.table_dir = os.path.join(self.base_dir, 'Medical', 'Table')
        self.excel_file = os.path.join(self.base_dir, 'Medical_Table_Results.xlsx')
        
        # 디렉토리 생성
        os.makedirs(self.table_dir, exist_ok=True)

    def pdf_to_png_memory(self, pdf_path, dpi=300):
        """PDF를 PNG로 변환 (메모리에서만 처리, 파일로 저장 안함)"""
        try:
            pdf_document = fitz.open(pdf_path)
            images = []
            
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                
                # 고해상도로 PNG 변환
                matrix = fitz.Matrix(dpi/72, dpi/72)
                pix = page.get_pixmap(matrix=matrix)
                
                # PIL Image로 변환
                img_data = pix.tobytes("pil")
                pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_data)
                
                # OpenCV 이미지로 변환
                cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
                
                images.append({
                    'page_num': page_num,
                    'image': cv_image,
                    'width': pix.width,
                    'height': pix.height
                })
            
            pdf_document.close()
            return images
            
        except Exception as e:
            print(f"PDF를 PNG로 변환 실패: {e}")
            return []

    def detect_table_regions(self, cv_image, min_area=5000):
        """이미지에서 테이블 영역 감지"""
        try:
            # 그레이스케일 변환
            gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
            
            # 이진화
            _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)
            
            # 수평선 감지
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
            horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
            
            # 수직선 감지
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
            vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
            
            # 수평선과 수직선 결합
            table_mask = cv2.addWeighted(horizontal_lines, 0.5, vertical_lines, 0.5, 0.0)
            
            # 노이즈 제거
            table_mask = cv2.morphologyEx(table_mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
            
            # 컨투어 찾기
            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 테이블 영역 후보 필터링
            table_regions = []
            for contour in contours:
                area = cv2.contourArea(contour)
                if area > min_area:  # 최소 면적 필터
                    x, y, w, h = cv2.boundingRect(contour)
                    
                    # 종횡비 체크 (너무 세로로 긴 것 제외)
                    aspect_ratio = w / h
                    if 0.3 < aspect_ratio < 10:
                        table_regions.append({
                            'x': x,
                            'y': y,
                            'width': w,
                            'height': h,
                            'area': area
                        })
            
            # 면적 기준으로 정렬
            table_regions.sort(key=lambda r: r['area'], reverse=True)
            
            return table_regions
            
        except Exception as e:
            print(f"테이블 영역 감지 실패: {e}")
            return []

    def extract_table_from_region(self, cv_image, region, padding=20):
        """특정 영역에서 테이블 이미지 추출"""
        try:
            # 패딩 추가
            x = max(0, region['x'] - padding)
            y = max(0, region['y'] - padding)
            x2 = min(cv_image.shape[1], region['x'] + region['width'] + padding)
            y2 = min(cv_image.shape[0], region['y'] + region['height'] + padding)
            
            # 테이블 영역 잘라내기
            table_image = cv_image[y:y2, x:x2]
            
            return table_image, (x, y, x2-x, y2-y)
            
        except Exception as e:
            print(f"테이블 이미지 추출 실패: {e}")
            return None, None

    def extract_tables_from_pdf_image(self, pdf_path, origin_number):
        """PDF를 이미지로 변환 후 테이블 추출"""
        try:
            print(f"PDF 이미지 변환 후 테이블 추출 시작: {pdf_path}")
            
            # PDF를 PNG 이미지로 변환 (메모리에서만)
            page_images = self.pdf_to_png_memory(pdf_path, dpi=300)
            
            if not page_images:
                print("PDF를 이미지로 변환할 수 없습니다.")
                return []
            
            table_info = []
            
            for page_data in page_images:
                page_num = page_data['page_num']
                cv_image = page_data['image']
                
                print(f"페이지 {page_num + 1} 처리 중... (크기: {page_data['width']}x{page_data['height']})")
                
                # 테이블 영역 감지
                table_regions = self.detect_table_regions(cv_image)
                
                if table_regions:
                    print(f"페이지 {page_num + 1}에서 {len(table_regions)}개의 테이블 영역을 발견했습니다.")
                    
                    for table_idx, region in enumerate(table_regions):
                        try:
                            # 테이블 이미지 추출
                            table_image, final_region = self.extract_table_from_region(cv_image, region)
                            
                            if table_image is not None:
                                # 테이블 이미지 저장
                                table_filename = f"M_table_{origin_number}_{len(table_info)}.png"
                                table_path = os.path.join(self.table_dir, table_filename)
                                
                                # OpenCV 이미지를 PIL로 변환 후 저장
                                pil_image = Image.fromarray(cv2.cvtColor(table_image, cv2.COLOR_BGR2RGB))
                                pil_image.save(table_path, "PNG", quality=95)
                                
                                # 테이블 정보 기록
                                table_info.append({
                                    'table_number': len(table_info),
                                    'filename': table_path,
                                    'page_number': page_num + 1,
                                    'table_index_in_page': table_idx,
                                    'preview_text': f"Image-based table from Page {page_num + 1}",
                                    'rows': 0,  # 이미지 기반에서는 행 수 계산 어려움
                                    'columns': 0,  # 이미지 기반에서는 열 수 계산 어려움
                                    'size': f"Image-based",
                                    'image_size': f"{final_region[2]}x{final_region[3]}",
                                    'position': f"Page {page_num + 1}",
                                    'detection_method': 'image_based',
                                    'region_area': region['area']
                                })
                                
                                print(f"✅ 이미지 기반 테이블 추출 완료: {table_filename} (페이지 {page_num + 1}, 영역 {table_idx + 1})")
                            
                        except Exception as table_error:
                            print(f"❌ 페이지 {page_num + 1}의 테이블 {table_idx + 1} 추출 실패: {table_error}")
                            continue
                else:
                    print(f"페이지 {page_num + 1}에서 테이블을 찾을 수 없습니다.")
            
            print(f"총 {len(table_info)}개의 테이블을 이미지 기반으로 추출했습니다.")
            return table_info
            
        except Exception as e:
            print(f"PDF 이미지 기반 테이블 추출 실패: {e}")
            return []

    def reprocess_all_pdfs_image_based(self):
        """모든 PDF를 이미지 기반으로 재처리"""
        print("🖼️  PDF 이미지 기반 테이블 재추출 시작")
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
        
        total_tables_extracted = 0
        
        for idx, (_, row) in enumerate(pdf_entries.iterrows(), 1):
            origin_number = row['Origin Number']
            pdf_filename = row['URL'].replace('PDF_FILE: ', '').strip()
            pdf_path = os.path.join(self.origin_dir, f'M_origin_{origin_number}.pdf')
            
            print(f"\n진행상황: {idx}/{len(pdf_entries)}")
            print(f"{'='*50}")
            print(f"이미지 기반 처리 중: {pdf_filename}")
            print(f"Origin Number: {origin_number}")
            print(f"{'='*50}")
            
            if os.path.exists(pdf_path):
                # 이미지 기반 테이블 추출
                table_info = self.extract_tables_from_pdf_image(pdf_path, origin_number)
                total_tables_extracted += len(table_info)
                
                print(f"✅ PDF 이미지 기반 처리 완료: {len(table_info)}개 테이블 추출")
            else:
                print(f"❌ PDF 파일 없음: {pdf_path}")
        
        print(f"\n{'='*70}")
        print(f"🎉 모든 PDF 이미지 기반 처리 완료!")
        print(f"📊 총 {total_tables_extracted}개의 테이블이 이미지 기반으로 추출되었습니다.")
        print(f"💡 기존 테이블 이미지들이 이미지 인식 기반으로 덮어씌워졌습니다.")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*70}")
        
        return True

def main():
    """메인 실행 함수"""
    try:
        extractor = PDFImageTableExtractor()
        
        print("🖼️  PDF를 이미지로 변환 후 테이블 영역을 감지하여 추출합니다.")
        print("💡 이미지 인식 기반으로 더 정확한 테이블 감지가 가능합니다.")
        print("⚠️  기존 테이블 이미지들이 새로운 버전으로 덮어씌워집니다.")
        print("\n계속하시겠습니까? (y/N): ", end="")
        
        try:
            user_input = input().strip().lower()
            if user_input not in ['y', 'yes', '예', 'ㅇ']:
                print("이미지 기반 처리가 취소되었습니다.")
                return False
        except KeyboardInterrupt:
            print("\n이미지 기반 처리가 취소되었습니다.")
            return False
        
        success = extractor.reprocess_all_pdfs_image_based()
        return success
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        return False

if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)