import os
import pandas as pd
import shutil
from datetime import datetime
import fitz  # PyMuPDF
from PIL import Image
import io

class PDFTableProcessor:
    def __init__(self, excel_filename="Medical_Table_Results.xlsx"):
        self.excel_filename = excel_filename
        self.temperal_pdf_dir = "temperal_pdf"
        self.target_origin_dir = "Medical/Context/Origin"
        self.target_table_dir = "Medical/Table"
        self.setup_directories()
        self.existing_data = self.load_existing_data()
        
    def setup_directories(self):
        """필요한 디렉토리 생성"""
        os.makedirs(self.target_origin_dir, exist_ok=True)
        os.makedirs(self.target_table_dir, exist_ok=True)
        os.makedirs(self.temperal_pdf_dir, exist_ok=True)
        print("디렉토리 설정 완료")
        
    def load_existing_data(self):
        """기존 엑셀 파일에서 데이터 로드"""
        try:
            if os.path.exists(self.excel_filename):
                main_df = pd.read_excel(self.excel_filename, sheet_name='Main Results')
                table_df = pd.read_excel(self.excel_filename, sheet_name='Table Details')
                
                print(f"기존 엑셀 파일 로드: {len(main_df)}개 URL/PDF 기록")
                
                # 기존에 처리된 PDF 파일명들 추출 (URL 컬럼에서)
                existing_pdfs = set()
                for url in main_df['URL'].tolist():
                    if isinstance(url, str) and url.startswith('PDF_FILE:'):
                        pdf_filename = url.replace('PDF_FILE: ', '').strip()
                        existing_pdfs.add(pdf_filename)
                
                return {
                    'main_data': main_df.to_dict('records'),
                    'table_data': table_df.to_dict('records'),
                    'existing_pdfs': existing_pdfs,
                    'max_origin_number': main_df['Origin Number'].max() if len(main_df) > 0 else -1
                }
            else:
                print("엑셀 파일이 없습니다. 새로 생성합니다.")
                return {
                    'main_data': [],
                    'table_data': [],
                    'existing_pdfs': set(),
                    'max_origin_number': -1
                }
                
        except Exception as e:
            print(f"기존 엑셀 파일 로드 실패: {e}")
            return {
                'main_data': [],
                'table_data': [],
                'existing_pdfs': set(),
                'max_origin_number': -1
            }
    
    def get_next_origin_number(self):
        """다음 Origin Number 반환"""
        return self.existing_data['max_origin_number'] + 1
    
    def find_pdf_files(self):
        """temperal_pdf 디렉토리에서 PDF 파일 찾기 및 중복 확인"""
        try:
            all_pdf_files = []
            new_pdf_files = []
            existing_pdfs = self.existing_data['existing_pdfs']
            
            for filename in os.listdir(self.temperal_pdf_dir):
                if filename.lower().endswith('.pdf'):
                    pdf_path = os.path.join(self.temperal_pdf_dir, filename)
                    all_pdf_files.append((filename, pdf_path))
            
            print(f"temperal_pdf에서 {len(all_pdf_files)}개의 PDF 파일을 발견했습니다.")
            
            # 중복 PDF 확인
            print(f"\n=== PDF 중복 검사 ===")
            print(f"기존 PDF 개수: {len(existing_pdfs)}")
            
            for filename, pdf_path in all_pdf_files:
                if filename in existing_pdfs:
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
            # PDF를 네이밍 컨벤션에 맞게 저장
            target_filename = f"M_origin_{origin_number}.pdf"
            target_path = os.path.join(self.target_origin_dir, target_filename)
            
            # PDF 복사
            shutil.copy2(pdf_path, target_path)
            print(f"PDF 저장: {target_path}")
            
            return target_path, None
            
        except Exception as e:
            print(f"PDF 이동 실패: {e}")
            return None, None
    
    def extract_tables_from_pdf(self, pdf_path, origin_number):
        """PDF를 HTML로 변환 후 Selenium으로 테이블 추출"""
        try:
            print(f"PDF HTML 변환 후 테이블 추출 시작: {pdf_path}")
            
            # PDF를 HTML로 변환 후 Selenium으로 처리
            table_info = self.pdf_to_html_with_selenium(pdf_path, origin_number)
            
            return table_info
            
        except Exception as e:
            print(f"PDF HTML 변환 후 테이블 추출 실패: {e}")
            return []
    
    def pdf_to_html_with_selenium(self, pdf_path, origin_number):
        """PDF를 HTML로 변환 후 Selenium으로 테이블 추출"""
        try:
            print(f"PDF를 HTML 변환 후 테이블 추출 시작: {pdf_path}")
            
            # PDF를 PNG로 변환 (메모리에서만, 저장 안함)
            pdf_document = fitz.open(pdf_path)
            
            # 임시 HTML 생성
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body { margin: 0; padding: 20px; }
                    .page { margin-bottom: 50px; }
                    img { max-width: 100%; height: auto; }
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
                import base64
                img_data = pix.tobytes("png")
                img_base64 = base64.b64encode(img_data).decode()
                
                html_content += f"""
                <div class="page" id="page_{page_num}">
                    <h3>Page {page_num + 1}</h3>
                    <img src="data:image/png;base64,{img_base64}" alt="Page {page_num + 1}">
                </div>
                """
            
            html_content += """
            </body>
            </html>
            """
            
            pdf_document.close()
            
            # 임시 HTML 파일 생성
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            # Selenium으로 HTML 로드하고 테이블 감지
            table_info = self.extract_tables_with_selenium(temp_html_path, origin_number)
            
            # 임시 파일 삭제
            os.unlink(temp_html_path)
            
            return table_info
            
        except Exception as e:
            print(f"PDF HTML 변환 실패: {e}")
            return []
    
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
        
        try:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            return driver
        except Exception as e:
            print(f"WebDriver 설정 실패: {e}")
            return None
    
    def extract_tables_with_selenium(self, html_path, origin_number):
        """Selenium으로 HTML에서 테이블 추출"""
        from selenium.webdriver.common.by import By
        import time
        
        driver = self.setup_webdriver()
        if not driver:
            return []
        
        try:
            # HTML 파일 로드
            driver.get(f"file://{html_path}")
            time.sleep(3)
            
            print("테이블 검색 및 캡처 시작...")
            
            # 이미지에서 직접 테이블 영역을 찾는 대신, 전체 이미지를 처리
            pages = driver.find_elements(By.CLASS_NAME, "page")
            
            if not pages:
                print("페이지를 찾을 수 없습니다.")
                return []
            
            table_info = []
            
            for page_idx, page_div in enumerate(pages):
                try:
                    # 페이지 이미지 요소 찾기
                    img_element = page_div.find_element(By.TAG_NAME, "img")
                    
                    if not img_element.is_displayed():
                        continue
                    
                    # 페이지로 스크롤
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", page_div)
                    time.sleep(2)
                    
                    # 전체 페이지 이미지를 테이블로 간주하여 저장
                    table_filename = f"M_table_{origin_number}_{page_idx}.png"
                    table_path = os.path.join(self.target_table_dir, table_filename)
                    
                    # 이미지 스크린샷
                    img_element.screenshot(table_path)
                    
                    # 테이블 정보 기록
                    table_info.append({
                        'table_number': page_idx,
                        'filename': table_path,
                        'preview_text': f"PDF Page {page_idx + 1} converted to image",
                        'rows': 0,
                        'columns': 0,
                        'size': "PDF_IMAGE",
                        'image_size': f"{img_element.size['width']}x{img_element.size['height']}",
                        'position': f"Page {page_idx + 1}",
                        'extraction_method': 'pdf_to_image'
                    })
                    
                    print(f"✅ PDF 페이지 이미지 추출 완료: {table_filename}")
                    
                except Exception as page_error:
                    print(f"❌ 페이지 {page_idx + 1} 처리 실패: {page_error}")
                    continue
            
            return table_info
            
        except Exception as e:
            print(f"Selenium 처리 실패: {e}")
            return []
        finally:
            if driver:
                driver.quit()
    
    def old_extract_method_backup(self, pdf_path, origin_number):
        """기존 PyMuPDF 방식 (백업용)"""
        try:
            pdf_document = fitz.open(pdf_path)
            table_info = []
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
                            
                            # 테이블 이미지 저장
                            table_filename = f"M_table_{origin_number}_{len(table_info)}.png"
                            table_path = os.path.join(self.target_table_dir, table_filename)
                            pix.save(table_path)
                            
                            # 테이블 데이터 추출
                            table_data = table.extract()
                            
                            # 테이블 텍스트 미리보기 생성
                            preview_text = ""
                            if table_data and len(table_data) > 0:
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
                                'size': f"{len(table_data)}x{len(table_data[0]) if table_data and len(table_data) > 0 else 0}",
                                'image_size': f"{int((expanded_rect.x1 - expanded_rect.x0) * 400/72)}x{int((expanded_rect.y1 - expanded_rect.y0) * 400/72)}",
                                'position': f"Page {page_num + 1}"
                            })
                            
                            print(f"테이블 추출 완료: {table_filename} (페이지 {page_num + 1})")
                            
                        except Exception as table_error:
                            print(f"페이지 {page_num + 1}의 테이블 {table_idx} 추출 실패: {table_error}")
                            continue
            
            pdf_document.close()
            print(f"총 {len(table_info)}개의 테이블을 추출했습니다.")
            return table_info
            
        except Exception as e:
            print(f"PDF 테이블 추출 실패: {e}")
            return []
    
    def process_pdf(self, pdf_filename, pdf_path, origin_number):
        """단일 PDF 파일 처리"""
        try:
            print(f"\n{'='*50}")
            print(f"처리 중: {pdf_filename}")
            print(f"Origin Number: {origin_number}")
            print(f"{'='*50}")
            
            # PDF를 Origin 디렉토리로 이동
            pdf_target_path, _ = self.move_pdf_to_origin(pdf_path, origin_number)
            if not pdf_target_path:
                return None
            
            # 테이블 추출
            table_info = self.extract_tables_from_pdf(pdf_path, origin_number)
            
            # 결과 정리
            result = {
                'origin_number': origin_number,
                'url': f"PDF_FILE: {pdf_filename}",  # PDF 파일명을 URL 위치에 저장
                'page_title': pdf_filename.replace('.pdf', ''),
                'png_filename': f"M_origin_{origin_number}.pdf",  # PDF 파일명으로 변경
                'pdf_filename': pdf_target_path,
                'table_count': len(table_info),
                'table_info': table_info,
                'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'source_type': 'PDF',
                'window_size': 'N/A (PDF)'
            }
            
            print(f"PDF 처리 완료: {len(table_info)}개 테이블 추출")
            return result
            
        except Exception as e:
            print(f"PDF 처리 실패 ({pdf_filename}): {e}")
            return None
    
    def update_excel_data(self, new_results):
        """기존 엑셀 데이터에 새로운 결과 추가"""
        # 메인 데이터 업데이트
        for result in new_results:
            if result:
                main_entry = {
                    'Origin Number': result['origin_number'],
                    'URL': result['url'],
                    'Page Title': result['page_title'],
                    'PNG Filename': result['png_filename'],
                    'Table Count': result['table_count'],
                    'Processing Time': result['processing_time'],
                    'User Agent': 'PDF_PROCESSOR',
                    'Window Size': result.get('window_size', 'N/A'),
                    'Source Type': result.get('source_type', 'PDF'),
                    'PDF Filename': result.get('pdf_filename', '')
                }
                self.existing_data['main_data'].append(main_entry)
                
                # 테이블 데이터 업데이트
                for table in result['table_info']:
                    table_entry = {
                        'Origin Number': result['origin_number'],
                        'URL': result['url'],
                        'Table Number': table['table_number'],
                        'Table Filename': table['filename'],
                        'Table Size (Rows x Cols)': table['size'],
                        'Image Size (Width x Height)': table['image_size'],
                        'Position (X, Y)': table['position'],
                        'Page Number': table.get('page_number', 'N/A'),
                        'Rows': table['rows'],
                        'Columns': table['columns'],
                        'Preview Text': table['preview_text']
                    }
                    self.existing_data['table_data'].append(table_entry)
                
                # 처리된 PDF를 기존 PDF 세트에 추가
                if result['url'].startswith('PDF_FILE:'):
                    pdf_filename = result['url'].replace('PDF_FILE: ', '').strip()
                    self.existing_data['existing_pdfs'].add(pdf_filename)
                
                # 최대 Origin Number 업데이트
                if result['origin_number'] > self.existing_data['max_origin_number']:
                    self.existing_data['max_origin_number'] = result['origin_number']
    
    def save_to_excel(self):
        """전체 데이터를 엑셀 파일로 저장"""
        try:
            print(f"\n엑셀 파일 업데이트 중: {self.excel_filename}")
            
            # 엑셀 파일 작성
            with pd.ExcelWriter(self.excel_filename, engine='openpyxl') as writer:
                # 메인 결과 시트
                main_df = pd.DataFrame(self.existing_data['main_data'])
                main_df.to_excel(writer, sheet_name='Main Results', index=False)
                
                # 테이블 상세 시트
                table_df = pd.DataFrame(self.existing_data['table_data'])
                table_df.to_excel(writer, sheet_name='Table Details', index=False)
            
            print(f"엑셀 파일 저장 완료: {self.excel_filename}")
            
            # 실제 파일 개수와 엑셀 기록 개수 비교
            actual_file_count = 0
            try:
                table_files = [f for f in os.listdir(self.target_table_dir) if f.endswith('.png')]
                actual_file_count = len(table_files)
            except Exception as e:
                print(f"실제 파일 개수 확인 실패: {e}")
            
            # 결과 요약
            total_entries = len(self.existing_data['main_data'])
            total_tables_in_excel = len(self.existing_data['table_data'])
            
            print(f"\n{'='*60}")
            print(f"전체 데이터베이스 현황 (PDF 처리 후)")
            print(f"{'='*60}")
            print(f"총 처리된 항목: {total_entries}개 (URL + PDF)")
            print(f"엑셀에 기록된 테이블: {total_tables_in_excel}개")
            print(f"실제 저장된 파일: {actual_file_count}개")
            if total_tables_in_excel != actual_file_count:
                hidden_tables = total_tables_in_excel - actual_file_count
                print(f"숨겨진/건너뛴 테이블: {hidden_tables}개")
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
        print("PDF 테이블 처리 시작")
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # temperal_pdf 디렉토리에서 새로운 PDF 파일 찾기
        pdf_files = self.find_pdf_files()
        
        if not pdf_files:
            print("처리할 새로운 PDF 파일이 없습니다. 모든 PDF가 이미 처리되었거나 파일이 없습니다.")
            
            # 현재 상태 표시
            actual_file_count = 0
            try:
                table_files = [f for f in os.listdir(self.target_table_dir) if f.endswith('.png')]
                actual_file_count = len(table_files)
                print(f"\n📁 테이블 디렉토리 파일 개수: {actual_file_count}개")
                print(f"디렉토리 경로: {self.target_table_dir}")
            except Exception as e:
                print(f"테이블 디렉토리 파일 개수 확인 실패: {e}")
            
            return
        
        print(f"총 {len(pdf_files)}개의 PDF 파일을 처리합니다.")
        
        # 새로운 결과 저장용
        new_results = []
        
        # 각 PDF 파일 처리
        for i, (pdf_filename, pdf_path) in enumerate(pdf_files):
            print(f"\n진행상황: {i+1}/{len(pdf_files)}")
            
            # Origin Number 계산
            origin_number = self.get_next_origin_number()
            self.existing_data['max_origin_number'] = origin_number  # 즉시 업데이트
            
            result = self.process_pdf(pdf_filename, pdf_path, origin_number)
            new_results.append(result)
            
            # 처리 결과를 즉시 엑셀에 저장 (중간 저장)
            if result:
                self.update_excel_data([result])
                self.save_to_excel()
                print(f"중간 저장 완료 (Origin {origin_number})")
                
                # 처리 완료된 PDF 파일 삭제
                self.cleanup_temperal_pdf(pdf_path)
            
            # 다음 PDF 처리 전 잠시 대기
            if i < len(pdf_files) - 1:
                print("다음 PDF 처리를 위해 1초 대기...")
                import time
                time.sleep(1)
        
        # 최종 저장 (이미 중간에 저장되었지만 확인차 한 번 더)
        if any(new_results):
            print("최종 엑셀 파일 저장 확인...")
            self.save_to_excel()
        
        print(f"\n모든 PDF 처리가 완료되었습니다!")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    processor = PDFTableProcessor()
    processor.run()