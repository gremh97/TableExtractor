#!/usr/bin/env python3
"""
PDF 처리기 - pdfplumber 기반 테이블 추출
pdfplumber와 pdf2image를 사용하여 정확한 테이블 영역만 추출
"""

import os
import sys
import shutil
import fitz  # PyMuPDF
import pandas as pd
import time
from datetime import datetime
import tempfile
import base64

class PDFTableProcessorPdfplumber:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.temperal_pdf_dir = os.path.join(self.base_dir, 'temperal_pdf')
        self.target_origin_dir = os.path.join(self.base_dir, 'Medical', 'Context', 'Origin')
        self.target_table_dir = os.path.join(self.base_dir, 'Medical', 'Table')
        self.excel_filename = os.path.join(self.base_dir, 'Medical_Table_Results.xlsx')
        
        # 디렉토리 생성
        for dir_path in [self.target_origin_dir, self.target_table_dir]:
            os.makedirs(dir_path, exist_ok=True)
        
        # 기존 데이터 로드
        self.existing_data = self.load_existing_data()

    def load_existing_data(self):
        """기존 엑셀 파일에서 데이터 로드"""
        try:
            if os.path.exists(self.excel_filename):
                main_df = pd.read_excel(self.excel_filename, sheet_name='Main Results')
                table_df = pd.read_excel(self.excel_filename, sheet_name='Table Details')
                
                # 기존 PDF 파일명 추출 (중복 방지용)
                existing_pdfs = set()
                for _, row in main_df.iterrows():
                    url_field = str(row.get('URL', ''))
                    if url_field.startswith('PDF_FILE:'):
                        pdf_filename = url_field.replace('PDF_FILE: ', '').strip()
                        existing_pdfs.add(pdf_filename)
                
                print(f"기존 엑셀 파일 로드: {len(main_df)}개 URL/PDF 기록")
                
                return {
                    'main_data': main_df.to_dict('records'),
                    'table_data': table_df.to_dict('records'),
                    'existing_pdfs': existing_pdfs,
                    'max_origin_number': int(main_df['Origin Number'].max()) if len(main_df) > 0 else 0
                }
            else:
                print("새로운 엑셀 파일을 생성합니다.")
                return {
                    'main_data': [],
                    'table_data': [],
                    'existing_pdfs': set(),
                    'max_origin_number': 0
                }
        except Exception as e:
            print(f"기존 데이터 로드 실패: {e}")
            return {
                'main_data': [],
                'table_data': [],
                'existing_pdfs': set(),
                'max_origin_number': 0
            }

    def find_pdf_files(self):
        """temperal_pdf에서 새로운 PDF 파일 찾기"""
        try:
            if not os.path.exists(self.temperal_pdf_dir):
                print(f"temperal_pdf 디렉토리가 없습니다: {self.temperal_pdf_dir}")
                return []
            
            all_pdf_files = []
            new_pdf_files = []
            
            for filename in os.listdir(self.temperal_pdf_dir):
                if filename.lower().endswith('.pdf'):
                    pdf_path = os.path.join(self.temperal_pdf_dir, filename)
                    all_pdf_files.append((filename, pdf_path))
            
            print(f"temperal_pdf에서 {len(all_pdf_files)}개의 PDF 파일을 발견했습니다.")
            
            # 중복 PDF 확인
            print(f"\n=== PDF 중복 검사 ===")
            print(f"기존 PDF 개수: {len(self.existing_data['existing_pdfs'])}")
            
            for filename, pdf_path in all_pdf_files:
                if filename in self.existing_data['existing_pdfs']:
                    print(f"중복 PDF (건너뜀): {filename}")
                else:
                    new_pdf_files.append((filename, pdf_path))
                    print(f"새로운 PDF (처리예정): {filename}")
            
            print(f"총 {len(new_pdf_files)}개의 새로운 PDF를 처리합니다.")
            return new_pdf_files
            
        except Exception as e:
            print(f"PDF 파일 검색 실패: {e}")
            return []

    def move_pdf_to_origin(self, pdf_path, origin_number):
        """PDF 파일을 Medical/Context/Origin으로 이동"""
        try:
            target_filename = f"M_origin_{origin_number}.pdf"
            target_path = os.path.join(self.target_origin_dir, target_filename)
            
            shutil.copy2(pdf_path, target_path)
            print(f"PDF 저장: {target_path}")
            
            return target_path
            
        except Exception as e:
            print(f"PDF 이동 실패: {e}")
            return None

    def setup_webdriver(self):
        """Chrome WebDriver 설정"""
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")
        
        try:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            return driver
        except Exception as e:
            print(f"WebDriver 설정 실패: {e}")
            return None

    def pdf_to_html(self, pdf_path):
        """PDF를 HTML로 변환 (이미지 포함)"""
        try:
            pdf_document = fitz.open(pdf_path)
            
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body { margin: 0; padding: 20px; background: white; }
                    .page { margin-bottom: 50px; border: 1px solid #ccc; padding: 10px; }
                    img { max-width: 100%; height: auto; display: block; }
                </style>
            </head>
            <body>
            """
            
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                
                # 고해상도로 PNG 변환
                matrix = fitz.Matrix(300/72, 300/72)  # 300 DPI
                pix = page.get_pixmap(matrix=matrix)
                
                # 이미지를 base64로 변환
                img_data = pix.tobytes("png")
                img_base64 = base64.b64encode(img_data).decode()
                
                html_content += f"""
                <div class="page" id="page_{page_num}">
                    <h3>Page {page_num + 1}</h3>
                    <img src="data:image/png;base64,{img_base64}" alt="Page {page_num + 1}" class="page-image">
                </div>
                """
            
            html_content += """
            </body>
            </html>
            """
            
            pdf_document.close()
            return html_content
            
        except Exception as e:
            print(f"PDF HTML 변환 실패: {e}")
            return None

    def detect_table_regions_opencv(self, image_path):
        """OpenCV를 사용해서 이미지에서 테이블 영역 감지"""
        import cv2
        import numpy as np
        
        try:
            # 이미지 읽기
            image = cv2.imread(image_path)
            if image is None:
                print(f"이미지를 읽을 수 없습니다: {image_path}")
                return []
            
            # 그레이스케일 변환
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # 이진화 (적응형 임계값 사용)
            binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 15, 10)
            
            # 수평선 감지를 위한 커널
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (50, 1))
            horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
            
            # 수직선 감지를 위한 커널
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))
            vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
            
            # 수평선과 수직선 결합
            table_structure = cv2.addWeighted(horizontal_lines, 0.5, vertical_lines, 0.5, 0.0)
            
            # 노이즈 제거와 구조 강화
            kernel = np.ones((3, 3), np.uint8)
            table_structure = cv2.morphologyEx(table_structure, cv2.MORPH_CLOSE, kernel)
            table_structure = cv2.dilate(table_structure, kernel, iterations=2)
            
            # 윤곽선 찾기
            contours, _ = cv2.findContours(table_structure, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 테이블 후보 영역 필터링
            table_regions = []
            min_area = 10000  # 최소 면적
            
            for contour in contours:
                area = cv2.contourArea(contour)
                if area > min_area:
                    x, y, w, h = cv2.boundingRect(contour)
                    
                    # 종횡비와 크기 조건 확인
                    aspect_ratio = w / h if h > 0 else 0
                    
                    # 테이블다운 영역 조건
                    if (w > 200 and h > 100 and  # 최소 크기
                        0.5 < aspect_ratio < 5.0 and  # 적절한 종횡비
                        w < image.shape[1] * 0.95 and  # 너무 크지 않음
                        h < image.shape[0] * 0.95):
                        
                        # 패딩 추가 (경계를 약간 넓게)
                        padding = 20
                        x = max(0, x - padding)
                        y = max(0, y - padding)
                        w = min(image.shape[1] - x, w + 2 * padding)
                        h = min(image.shape[0] - y, h + 2 * padding)
                        
                        table_regions.append({
                            'x': x,
                            'y': y,
                            'width': w,
                            'height': h,
                            'area': area,
                            'aspect_ratio': aspect_ratio
                        })
            
            # 면적 순으로 정렬 (큰 것부터)
            table_regions.sort(key=lambda r: r['area'], reverse=True)
            
            print(f"감지된 테이블 영역: {len(table_regions)}개")
            for i, region in enumerate(table_regions):
                print(f"  테이블 {i+1}: 위치({region['x']}, {region['y']}), 크기({region['width']}x{region['height']}), 면적({region['area']})")
            
            return table_regions
            
        except Exception as e:
            print(f"테이블 영역 감지 실패: {e}")
            return []

    def extract_table_region(self, image_path, region, output_path):
        """이미지에서 특정 테이블 영역 추출"""
        import cv2
        import numpy as np
        
        try:
            # 이미지 읽기
            image = cv2.imread(image_path)
            if image is None:
                return False
            
            # 테이블 영역 잘라내기
            x, y, w, h = region['x'], region['y'], region['width'], region['height']
            table_image = image[y:y+h, x:x+w]
            
            # 이미지 품질 향상
            # 선명도 향상
            kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
            table_image = cv2.filter2D(table_image, -1, kernel)
            
            # 저장
            cv2.imwrite(output_path, table_image, [cv2.IMWRITE_PNG_COMPRESSION, 0])
            return True
            
        except Exception as e:
            print(f"테이블 영역 추출 실패: {e}")
            return False

    def extract_tables_with_selenium(self, html_content, origin_number):
        """Selenium으로 HTML에서 테이블 추출 (OpenCV 테이블 감지)"""
        from selenium.webdriver.common.by import By
        import cv2
        import numpy as np
        
        driver = self.setup_webdriver()
        if not driver:
            return []
        
        try:
            # 임시 HTML 파일 생성
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            # HTML 파일 로드
            driver.get(f"file://{temp_html_path}")
            time.sleep(3)
            
            print("페이지 이미지에서 테이블 감지 시작...")
            
            # 페이지 이미지들 찾기
            page_images = driver.find_elements(By.CLASS_NAME, "page-image")
            
            if not page_images:
                print("페이지 이미지를 찾을 수 없습니다.")
                return []
            
            print(f"{len(page_images)}개의 페이지 이미지를 발견했습니다.")
            
            table_info = []
            
            for page_idx, img_element in enumerate(page_images):
                try:
                    if not img_element.is_displayed():
                        continue
                    
                    # 이미지로 스크롤
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", img_element)
                    time.sleep(2)
                    
                    # 이미지 크기 확인
                    size = img_element.size
                    if size['width'] < 100 or size['height'] < 100:
                        print(f"페이지 {page_idx + 1} 이미지가 너무 작아 건너뜁니다.")
                        continue
                    
                    # 임시로 전체 페이지 이미지 저장
                    temp_page_path = os.path.join(tempfile.gettempdir(), f"temp_page_{origin_number}_{page_idx}.png")
                    img_element.screenshot(temp_page_path)
                    
                    # OpenCV로 테이블 영역 감지
                    table_regions = self.detect_table_regions_opencv(temp_page_path)
                    
                    if table_regions:
                        # 각 테이블 영역별로 저장
                        for table_idx, region in enumerate(table_regions):
                            table_filename = f"M_table_{origin_number}_{len(table_info)}.png"
                            table_path = os.path.join(self.target_table_dir, table_filename)
                            
                            # 테이블 영역만 추출해서 저장
                            if self.extract_table_region(temp_page_path, region, table_path):
                                table_info.append({
                                    'table_number': len(table_info),
                                    'filename': table_path,
                                    'preview_text': f"PDF Page {page_idx + 1} Table {table_idx + 1} - OpenCV detected",
                                    'rows': 0,
                                    'columns': 0,
                                    'size': "OPENCV_TABLE",
                                    'image_size': f"{region['width']}x{region['height']}",
                                    'position': f"Page {page_idx + 1} Table {table_idx + 1}",
                                    'extraction_method': 'opencv_table_detection',
                                    'region_area': region['area'],
                                    'aspect_ratio': region['aspect_ratio']
                                })
                                
                                print(f"✅ 테이블 영역 저장 완료: {table_filename} (페이지 {page_idx + 1}, 테이블 {table_idx + 1})")
                            else:
                                print(f"❌ 테이블 영역 저장 실패: {table_filename}")
                    else:
                        # 테이블이 감지되지 않으면 전체 페이지를 저장 (기존 방식)
                        table_filename = f"M_table_{origin_number}_{len(table_info)}.png"
                        table_path = os.path.join(self.target_table_dir, table_filename)
                        
                        img_element.screenshot(table_path)
                        
                        table_info.append({
                            'table_number': len(table_info),
                            'filename': table_path,
                            'preview_text': f"PDF Page {page_idx + 1} - full page (no table detected)",
                            'rows': 0,
                            'columns': 0,
                            'size': "FULL_PAGE",
                            'image_size': f"{size['width']}x{size['height']}",
                            'position': f"Page {page_idx + 1}",
                            'extraction_method': 'full_page_fallback'
                        })
                        
                        print(f"⚠️ 테이블 미감지, 전체 페이지 저장: {table_filename}")
                    
                    # 임시 파일 삭제
                    if os.path.exists(temp_page_path):
                        os.unlink(temp_page_path)
                    
                except Exception as page_error:
                    print(f"❌ 페이지 {page_idx + 1} 처리 실패: {page_error}")
                    continue
            
            # 임시 HTML 파일 삭제
            os.unlink(temp_html_path)
            
            print(f"총 {len(table_info)}개의 테이블/이미지를 저장했습니다.")
            return table_info
            
        except Exception as e:
            print(f"Selenium 처리 실패: {e}")
            return []
        finally:
            if driver:
                driver.quit()

    def extract_tables_from_pdf_direct(self, pdf_path, origin_number):
        """pdfplumber와 pdf2image를 사용해서 정확한 테이블 영역만 감지하여 추출"""
        import pdfplumber
        from pdf2image import convert_from_path
        from PIL import Image
        
        try:
            print(f"PDF에서 테이블 영역 감지하여 추출: {pdf_path}")
            
            # 1. PDF 페이지를 고해상도 이미지로 변환
            print("PDF 페이지를 고해상도 이미지로 변환 중...")
            try:
                # 고해상도로 변환 (300 DPI)
                pages_as_images = convert_from_path(pdf_path, dpi=300)
                print(f"{len(pages_as_images)}개 페이지를 이미지로 변환 완료")
            except Exception as e:
                print(f"PDF 이미지 변환 실패 (Poppler 설치 확인): {e}")
                return []
            
            # 2. pdfplumber로 테이블 위치 감지
            table_info = []
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    print(f"페이지 {page_num + 1} 테이블 감지 중...")
                    
                    # pdfplumber로 테이블 찾기
                    try:
                        tables = page.find_tables()
                        
                        if tables:
                            print(f"페이지 {page_num + 1}에서 {len(tables)}개의 테이블을 발견했습니다.")
                            
                            # 해당 페이지의 이미지
                            page_image = pages_as_images[page_num]
                            
                            for table_idx, table in enumerate(tables):
                                try:
                                    # 테이블의 바운딩 박스 (x0, top, x1, bottom)
                                    bbox = table.bbox
                                    print(f"  원본 bbox: {bbox}")
                                    print(f"  페이지 크기: {page.width} x {page.height}")
                                    print(f"  이미지 크기: {page_image.width} x {page_image.height}")
                                    
                                    # PDF 포인트를 이미지 픽셀 좌표로 변환
                                    scale_x = page_image.width / page.width
                                    scale_y = page_image.height / page.height
                                    print(f"  스케일: {scale_x:.3f} x {scale_y:.3f}")
                                    
                                    # PIL crop 좌표 (left, top, right, bottom)
                                    left = int(bbox[0] * scale_x)
                                    top = int(bbox[1] * scale_y)
                                    right = int(bbox[2] * scale_x)
                                    bottom = int(bbox[3] * scale_y)
                                    print(f"  변환된 좌표: left={left}, top={top}, right={right}, bottom={bottom}")
                                    
                                    # 테이블 영역을 페이지 전체 너비로 확장 (오른쪽 잘림 완전 해결)
                                    original_left, original_top, original_right, original_bottom = left, top, right, bottom
                                    
                                    # 왼쪽은 약간 패딩, 오른쪽은 페이지 끝까지
                                    left = max(0, left - 30)
                                    top = max(0, top - 30) 
                                    right = page_image.width - 20  # 페이지 오른쪽 끝에서 20픽셀만 여백
                                    bottom = min(page_image.height, bottom + 30)
                                    
                                    print(f"  원본 테이블 영역: {original_left}~{original_right} (너비: {original_right-original_left})")
                                    print(f"  확장된 영역: {left}~{right} (너비: {right-left})")
                                    print(f"  확장된 크기: {right-left} x {bottom-top}")
                                    
                                    # 테이블 영역만 잘라내기
                                    cropped_table = page_image.crop((left, top, right, bottom))
                                    
                                    # 테이블 이미지 저장
                                    table_filename = f"M_table_{origin_number}_{len(table_info)}.png"
                                    table_path = os.path.join(self.target_table_dir, table_filename)
                                    cropped_table.save(table_path, "PNG")
                                    
                                    # 테이블 데이터 추출 시도
                                    try:
                                        table_data = table.extract()
                                        rows = len(table_data) if table_data else 0
                                        cols = len(table_data[0]) if table_data and len(table_data) > 0 else 0
                                        
                                        # 미리보기 텍스트 생성
                                        preview_text = ""
                                        if table_data and len(table_data) > 0:
                                            for row_idx, row in enumerate(table_data[:2]):  # 처음 2행만
                                                if row:
                                                    row_text = " | ".join([str(cell) if cell else "" for cell in row])
                                                    preview_text += row_text + " "
                                                    if len(preview_text) > 100:
                                                        break
                                            preview_text = preview_text.strip()[:150] + ("..." if len(preview_text) > 150 else "")
                                        
                                        if not preview_text:
                                            preview_text = f"Page {page_num + 1} Table {table_idx + 1}"
                                            
                                    except Exception as data_error:
                                        print(f"테이블 데이터 추출 실패: {data_error}")
                                        rows, cols = 0, 0
                                        preview_text = f"Page {page_num + 1} Table {table_idx + 1}"
                                    
                                    table_info.append({
                                        'table_number': len(table_info),
                                        'filename': table_path,
                                        'preview_text': preview_text,
                                        'rows': rows,
                                        'columns': cols,
                                        'size': f"{rows}x{cols}" if rows > 0 and cols > 0 else "DETECTED",
                                        'image_size': f"{cropped_table.width}x{cropped_table.height}",
                                        'position': f"Page {page_num + 1} Table {table_idx + 1}",
                                        'extraction_method': 'pdfplumber_table_detection'
                                    })
                                    
                                    print(f"✅ 테이블 영역 추출 완료: {table_filename} (페이지 {page_num + 1}, 테이블 {table_idx + 1}) - 크기: {cropped_table.width}x{cropped_table.height}")
                                    
                                except Exception as table_error:
                                    print(f"❌ 페이지 {page_num + 1}의 테이블 {table_idx + 1} 추출 실패: {table_error}")
                                    continue
                        
                        else:
                            # 테이블이 감지되지 않은 경우 - 건너뜀 (전체 페이지 저장하지 않음)
                            print(f"⚠️ 페이지 {page_num + 1}에서 테이블을 감지하지 못했습니다. (건너뛰기)")
                    
                    except Exception as page_error:
                        print(f"❌ 페이지 {page_num + 1} 처리 실패: {page_error}")
                        continue
            
            print(f"총 {len(table_info)}개의 테이블을 추출했습니다.")
            
            # 테이블이 하나도 없는 경우
            if len(table_info) == 0:
                print("⚠️ 감지된 테이블이 없습니다. 전체 페이지도 저장하지 않습니다.")
            
            return table_info
            
        except Exception as e:
            print(f"PDF 테이블 추출 실패: {e}")
            return []

    def process_single_pdf(self, pdf_filename, pdf_path):
        """단일 PDF 파일 처리"""
        try:
            # 다음 Origin Number 계산
            origin_number = self.existing_data['max_origin_number'] + 1
            
            print(f"\n{'='*50}")
            print(f"처리 중: {pdf_filename}")
            print(f"Origin Number: {origin_number}")
            print(f"{'='*50}")
            
            # PDF를 Origin 디렉토리로 이동
            pdf_target_path = self.move_pdf_to_origin(pdf_path, origin_number)
            if not pdf_target_path:
                return None
            
            # PyMuPDF로 실제 테이블 영역만 추출
            table_info = self.extract_tables_from_pdf_direct(pdf_path, origin_number)
            
            # 결과 정리
            result = {
                'origin_number': origin_number,
                'url': f"PDF_FILE: {pdf_filename}",
                'page_title': pdf_filename.replace('.pdf', ''),
                'png_filename': f"M_origin_{origin_number}.pdf",
                'pdf_filename': pdf_target_path,
                'table_count': len(table_info),
                'table_info': table_info,
                'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            
            print(f"PDF 처리 완료: {len(table_info)}개 페이지 이미지 추출")
            return result
            
        except Exception as e:
            print(f"PDF 처리 실패: {e}")
            return None

    def update_excel_data(self, result):
        """엑셀 데이터 업데이트"""
        try:
            # 메인 데이터 추가
            main_entry = {
                'Origin Number': result['origin_number'],
                'URL': result['url'],
                'Page Title': result['page_title'],
                'PNG Filename': result['png_filename'],
                'Table Count': result['table_count'],
                'Processing Time': result['processing_time']
            }
            self.existing_data['main_data'].append(main_entry)
            
            # 테이블 상세 데이터 추가
            for table_info in result['table_info']:
                table_entry = {
                    'Origin Number': result['origin_number'],
                    'URL': result['url'],
                    'Table Number': table_info['table_number'],
                    'Table Filename': os.path.basename(table_info['filename']),
                    'Preview Text': table_info['preview_text'],
                    'Rows': table_info['rows'],
                    'Columns': table_info['columns'],
                    'Size': table_info['size'],
                    'Image Size': table_info['image_size'],
                    'Position': table_info['position'],
                    'Extraction Method': table_info['extraction_method']
                }
                self.existing_data['table_data'].append(table_entry)
            
            # 처리된 PDF를 기존 PDF 세트에 추가
            if result['url'].startswith('PDF_FILE:'):
                pdf_filename = result['url'].replace('PDF_FILE: ', '').strip()
                self.existing_data['existing_pdfs'].add(pdf_filename)
            
            # 최대 Origin Number 업데이트
            if result['origin_number'] > self.existing_data['max_origin_number']:
                self.existing_data['max_origin_number'] = result['origin_number']
                
        except Exception as e:
            print(f"엑셀 데이터 업데이트 실패: {e}")

    def save_to_excel(self):
        """전체 데이터를 엑셀 파일로 저장"""
        try:
            print(f"\n엑셀 파일 업데이트 중: {self.excel_filename}")
            
            with pd.ExcelWriter(self.excel_filename, engine='openpyxl') as writer:
                # 메인 결과 시트
                main_df = pd.DataFrame(self.existing_data['main_data'])
                main_df.to_excel(writer, sheet_name='Main Results', index=False)
                
                # 테이블 상세 시트
                table_df = pd.DataFrame(self.existing_data['table_data'])
                table_df.to_excel(writer, sheet_name='Table Details', index=False)
            
            print(f"엑셀 파일 저장 완료: {self.excel_filename}")
            
            # 상태 출력
            print(f"\n{'='*60}")
            print(f"전체 데이터베이스 현황")
            print(f"{'='*60}")
            print(f"총 처리된 항목: {len(self.existing_data['main_data'])}개 (URL + PDF)")
            print(f"엑셀에 기록된 테이블: {len(self.existing_data['table_data'])}개")
            
            # 실제 파일 개수
            if os.path.exists(self.target_table_dir):
                actual_files = len([f for f in os.listdir(self.target_table_dir) if f.startswith('M_table_')])
                print(f"실제 저장된 파일: {actual_files}개")
            
            print(f"최대 Origin Number: {self.existing_data['max_origin_number']}")
            print(f"엑셀 파일: {self.excel_filename}")
            print(f"PDF 저장 위치: {self.target_origin_dir}/")
            print(f"테이블 이미지 저장 위치: {self.target_table_dir}/")
            print(f"{'='*60}")
            
        except Exception as e:
            print(f"엑셀 저장 실패: {e}")

    def cleanup_temperal_pdf(self, pdf_path):
        """처리 완료된 PDF 파일을 temperal_pdf에서 제거"""
        try:
            os.remove(pdf_path)
            print(f"처리 완료된 PDF 파일 삭제: {pdf_path}")
        except Exception as e:
            print(f"PDF 파일 삭제 실패: {e}")

    def run(self):
        """메인 실행 함수"""
        print("PDF 테이블 처리 시작 (Selenium 기반)")
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # temperal_pdf 디렉토리에서 새로운 PDF 파일 찾기
        pdf_files = self.find_pdf_files()
        
        if not pdf_files:
            print("처리할 새로운 PDF 파일이 없습니다.")
            self.save_to_excel()  # 현재 상태 표시
            return
        
        print(f"총 {len(pdf_files)}개의 PDF 파일을 처리합니다.\n")
        
        # 각 PDF 파일 처리
        for idx, (pdf_filename, pdf_path) in enumerate(pdf_files, 1):
            print(f"진행상황: {idx}/{len(pdf_files)}")
            
            # PDF 처리
            result = self.process_single_pdf(pdf_filename, pdf_path)
            
            if result:
                # 엑셀 데이터 업데이트
                self.update_excel_data(result)
                
                # 중간 저장
                self.save_to_excel()
                print(f"중간 저장 완료 (Origin {result['origin_number']})")
                
                # 처리 완료된 PDF 삭제
                self.cleanup_temperal_pdf(pdf_path)
                
                if idx < len(pdf_files):
                    print("다음 PDF 처리를 위해 1초 대기...")
                    time.sleep(1)
        
        # 최종 저장
        print("\n최종 엑셀 파일 저장 확인...")
        self.save_to_excel()
        
        print(f"\n모든 PDF 처리가 완료되었습니다!")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

def main():
    """프로그램 진입점"""
    try:
        print("디렉토리 설정 완료")
        
        processor = PDFTableProcessorPdfplumber()
        processor.run()
        
    except KeyboardInterrupt:
        print("\n\n사용자가 프로그램을 중단했습니다.")
        sys.exit(1)
    except Exception as e:
        print(f"\n예상치 못한 오류가 발생했습니다: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()