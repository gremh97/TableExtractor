import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from PIL import Image
import io
import requests
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import numpy as np
import urllib3
import tempfile
import base64
import ssl
from urllib.parse import urlparse
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class ContinuousPNGTableExtractor:
    def __init__(self, excel_filename="Medical_Table_Results.xlsx"):
        self.excel_filename = excel_filename
        self.setup_directories()
        self.existing_data = self.load_existing_data()
        
    def setup_directories(self):
        """필요한 디렉토리 생성"""
        os.makedirs("Medical/Context/Origin", exist_ok=True)
        os.makedirs("Medical/Table", exist_ok=True)
        print("디렉토리 설정 완료")
        
    def load_existing_data(self):
        """기존 엑셀 파일에서 데이터 로드"""
        try:
            if os.path.exists(self.excel_filename):
                main_df = pd.read_excel(self.excel_filename, sheet_name='Main Results')
                table_df = pd.read_excel(self.excel_filename, sheet_name='Table Details')
                
                print(f"기존 엑셀 파일 로드: {len(main_df)}개 URL 기록")
                
                return {
                    'main_data': main_df.to_dict('records'),
                    'table_data': table_df.to_dict('records'),
                    'existing_urls': set(main_df['URL'].tolist()),
                    'max_origin_number': main_df['Origin Number'].max() if len(main_df) > 0 else -1
                }
            else:
                print("새로운 엑셀 파일을 생성합니다.")
                return {
                    'main_data': [],
                    'table_data': [],
                    'existing_urls': set(),
                    'max_origin_number': -1
                }
                
        except Exception as e:
            print(f"기존 엑셀 파일 로드 실패: {e}")
            return {
                'main_data': [],
                'table_data': [],
                'existing_urls': set(),
                'max_origin_number': -1
            }
    
    def get_next_origin_number(self):
        """다음 Origin Number 반환"""
        return self.existing_data['max_origin_number'] + 1
    
    def filter_new_urls(self, urls):
        """중복되지 않는 새로운 URL만 필터링"""
        new_urls = []
        existing_urls = self.existing_data['existing_urls']
        
        print(f"\n=== URL 중복 검사 ===")
        print(f"기존 URL 개수: {len(existing_urls)}")
        
        for url in urls:
            if url in existing_urls:
                print(f"중복 URL (건너뜀): {url}")
            else:
                new_urls.append(url)
                print(f"새로운 URL (처리예정): {url}")
        
        print(f"총 {len(new_urls)}개의 새로운 URL을 처리합니다.")
        return new_urls
        
    def setup_webdriver(self):
        """Chrome WebDriver 설정 - 데스크톱 버전 강제"""
        chrome_options = Options()
        
        # 데스크톱 버전 강제 설정
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--force-device-scale-factor=1")
        
        # 모바일 에뮬레이션 비활성화
        chrome_options.add_argument("--disable-mobile-emulation")
        
        # 기본 설정
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        
        # 헤드리스 모드
        chrome_options.add_argument("--headless")
        
        # 실험적 옵션으로 데스크톱 강제
        chrome_options.add_experimental_option("useAutomationExtension", False)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        try:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # 윈도우 크기 명시적 설정
            driver.set_window_size(1920, 1080)
            
            return driver
        except Exception as e:
            print(f"WebDriver 설정 실패: {e}")
            return None
        
    def read_urls(self, filename="urls.txt"):
        """URL 파일 읽기"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                urls = [line.strip() for line in f if line.strip()]
            print(f"URL 파일 읽기 완료: {len(urls)}개 URL")
            return urls
        except FileNotFoundError:
            print(f"URL 파일 '{filename}'을 찾을 수 없습니다.")
            return []
    
    def scroll_page_completely(self, driver):
        """페이지 전체를 천천히 스크롤하여 모든 콘텐츠 로드"""
        print("페이지 스크롤 시작...")
        
        # 페이지 상단으로 이동
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(3)
        
        # 페이지 높이 가져오기
        last_height = driver.execute_script("return document.body.scrollHeight")
        scroll_position = 0
        
        # 천천히 스크롤하면서 콘텐츠 로드
        while scroll_position < last_height:
            # 현재 위치에서 500px씩 스크롤 (속도 향상)
            scroll_position += 500
            driver.execute_script(f"window.scrollTo(0, {scroll_position});")
            time.sleep(0.2)
            
            # 페이지 높이 다시 확인 (동적 콘텐츠 로딩)
            current_height = driver.execute_script("return document.body.scrollHeight")
            if current_height > last_height:
                last_height = current_height
        
        # 페이지 맨 끝까지 스크롤
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        
        # 페이지 상단으로 돌아가기
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        print("페이지 스크롤 완료")
    
    def save_page_as_png(self, driver, url, png_filename):
        """웹페이지를 PNG로 저장 (전체 페이지)"""
        try:
            print(f"PNG 저장 시작: {png_filename}")
            
            # 페이지 로딩 대기
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 추가 로딩 대기
            time.sleep(5)
            
            # 페이지 전체 스크롤
            self.scroll_page_completely(driver)
            
            # 전체 페이지 높이와 너비 가져오기
            total_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
            total_width = driver.execute_script("return Math.max( document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth );")
            
            print(f"페이지 크기: {total_width} x {total_height}")
            
            # 윈도우 크기를 페이지 크기에 맞게 조정
            driver.set_window_size(total_width, total_height)
            time.sleep(2)
            
            # 페이지 상단으로 이동
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)
            
            # 전체 페이지 스크린샷
            screenshot = driver.get_screenshot_as_png()
            
            # PNG 파일 저장
            with open(png_filename, 'wb') as f:
                f.write(screenshot)
            
            print(f"PNG 저장 완료: {png_filename}")
            return True
            
        except Exception as e:
            print(f"PNG 저장 실패: {e}")
            return False
    
    def capture_tables_as_images(self, driver, origin_number):
        """페이지의 테이블들을 이미지로 캡처"""
        try:
            print("테이블 검색 및 캡처 시작...")
            
            # 모든 테이블 요소 찾기
            tables = driver.find_elements(By.TAG_NAME, "table")
            
            if not tables:
                print("테이블을 찾을 수 없습니다.")
                return []
            
            print(f"{len(tables)}개의 테이블을 발견했습니다.")
            
            table_info = []
            
            for i, table in enumerate(tables):
                try:
                    # 테이블이 보이는지 확인
                    if not table.is_displayed():
                        print(f"테이블 {i}이 숨겨져 있어 건너뜁니다. (엑셀 기록 제외)")
                        continue
                    
                    # 테이블이 화면에 보이도록 스크롤
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", table)
                    time.sleep(2)
                    
                    # 테이블 크기 확인
                    size = table.size
                    location = table.location
                    
                    print(f"테이블 {i} 정보: 위치({location['x']}, {location['y']}), 크기({size['width']}x{size['height']})")
                    
                    if size['width'] < 50 or size['height'] < 50:
                        print(f"테이블 {i}이 너무 작아 건너뜁니다. (엑셀 기록 제외)")
                        continue
                    
                    # 테이블 스크린샷 촬영
                    table_filename = f"Medical/Table/M_table_{origin_number}_{i}.png"
                    table.screenshot(table_filename)
                    
                    # 테이블 정보 수집
                    try:
                        # 텍스트 추출을 더 안전하게
                        try:
                            table_text = table.text
                            if not table_text or table_text.strip() == "":
                                table_text = "텍스트 없음"
                            else:
                                table_text = table_text[:200].replace('\n', ' ').strip()
                        except:
                            table_text = "텍스트 추출 실패"
                        
                        # 테이블 행/열 수 계산
                        try:
                            rows = len(table.find_elements(By.TAG_NAME, "tr"))
                            if rows > 0:
                                first_row_elements = table.find_elements(By.TAG_NAME, "tr")
                                if first_row_elements:
                                    first_row = first_row_elements[0]
                                    th_elements = first_row.find_elements(By.TAG_NAME, "th")
                                    td_elements = first_row.find_elements(By.TAG_NAME, "td")
                                    cols = len(th_elements) + len(td_elements)
                                else:
                                    cols = 0
                            else:
                                cols = 0
                        except Exception as row_error:
                            print(f"테이블 {i} 행/열 계산 오류: {row_error}")
                            rows, cols = 0, 0
                        
                    except Exception as text_error:
                        print(f"테이블 {i} 정보 추출 오류: {text_error}")
                        table_text = f"Table {i} (정보 추출 실패)"
                        rows, cols = 0, 0
                    
                    table_info.append({
                        'table_number': i,
                        'filename': table_filename,
                        'preview_text': table_text,
                        'rows': rows,
                        'columns': cols,
                        'size': f"{rows}x{cols}",
                        'image_size': f"{size['width']}x{size['height']}",
                        'position': f"({location['x']}, {location['y']})"
                    })
                    
                    print(f"테이블 {i} 캡처 완료: {table_filename}")
                    
                except Exception as e:
                    print(f"테이블 {i} 캡처 실패: {e}")
                    continue
            
            print(f"총 {len(table_info)}개의 테이블 이미지 저장 완료")
            return table_info
            
        except Exception as e:
            print(f"테이블 캡처 중 오류 발생: {e}")
            return []

    def render_html_table_as_image(self, table_html, table_counter, origin_number):
        """HTML 테이블을 웹브라우저처럼 렌더링하여 이미지로 캡처"""
        try:
            # Chrome 옵션 설정
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1200,800')
            
            # 한글 폰트 지원을 위한 설정
            chrome_options.add_argument('--font-render-hinting=none')
            chrome_options.add_argument('--disable-font-subpixel-positioning')
            
            # WebDriver 초기화
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # 스타일이 포함된 HTML 생성
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif;
                        margin: 20px;
                        background-color: white;
                    }}
                    table {{
                        border-collapse: collapse;
                        width: 100%;
                        margin: 10px 0;
                        font-size: 14px;
                    }}
                    th, td {{
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                        vertical-align: top;
                    }}
                    th {{
                        background-color: #f2f2f2;
                        font-weight: bold;
                    }}
                    tr:nth-child(even) {{
                        background-color: #f9f9f9;
                    }}
                    .panel {{
                        border: 1px solid #ccc;
                        padding: 15px;
                        margin: 10px 0;
                        background-color: #fafafa;
                        border-radius: 5px;
                    }}
                </style>
            </head>
            <body>
                <div class="panel">
                    {table_html}
                </div>
            </body>
            </html>
            """
            
            # 임시 HTML 파일 생성
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            try:
                # HTML 파일 로드
                driver.get(f'file://{temp_html_path}')
                time.sleep(2)  # 렌더링 대기
                
                # 테이블 요소 찾기 및 캡처
                table_element = driver.find_element(By.TAG_NAME, 'table')
                
                # PNG 파일명
                png_filename = f"Medical/Table/M_table_{origin_number}_{table_counter}.png"
                os.makedirs(os.path.dirname(png_filename), exist_ok=True)
                
                # 스크린샷 저장
                table_element.screenshot(png_filename)
                
                return png_filename
                
            finally:
                driver.quit()
                # 임시 파일 삭제
                try:
                    os.unlink(temp_html_path)
                except:
                    pass
                    
        except Exception as e:
            print(f"HTML 테이블 렌더링 실패: {e}")
            return None

    def extract_hidden_tables_from_url(self, url, origin_number):
        """URL에서 HTML 직접 다운로드하여 panel 블록의 테이블 추출"""
        try:
            print(f"HTML 직접 다운로드 및 테이블 추출: {url}")
            
            # SSL 인증 우회 설정
            ssl_context = ssl.create_default_context()
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE
            
            # requests로 HTML 다운로드
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(url, headers=headers, verify=False, timeout=30)
            response.raise_for_status()
            response.encoding = 'utf-8'
            
            # BeautifulSoup으로 파싱
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # davoshospital.co.kr에서만 전체 HTML 저장
            if 'davoshospital.co.kr' in url:
                html_filename = f"Medical/Context/Origin/M_origin_{origin_number}.html"
                os.makedirs(os.path.dirname(html_filename), exist_ok=True)
                with open(html_filename, 'w', encoding='utf-8') as f:
                    f.write(response.text)
                print(f"전체 HTML 페이지 저장: {html_filename}")
            
            # panel 블록 찾기
            panels = soup.find_all('div', class_='panel')
            print(f"발견된 panel 블록 수: {len(panels)}")

            table_info = []
            table_counter = 0

            for p_idx, panel in enumerate(panels):
                # panel 내부의 모든 table 태그
                tables = panel.find_all('table')
                print(f" panel {p_idx}: 테이블 {len(tables)}개 발견")

                for t_idx, table in enumerate(tables):
                    # 판다스로 테이블 파싱 시도 (PNG 생성용)
                    try:
                        from io import StringIO
                        dfs = pd.read_html(StringIO(str(table)))
                    except Exception as e:
                        print(f"pd.read_html 실패: {e}")
                        dfs = []

                    if not dfs:
                        print(f"테이블 {table_counter}에 파싱 가능한 데이터가 없습니다. 건너뜀니다.")
                        continue

                    df = dfs[0]
                    
                    # 테이블 HTML은 이미 전체 페이지로 저장됨

                    # 저장: PNG (웹브라우저 스타일 렌더링)
                    print(f"HTML 테이블 렌더링 시도 중: 테이블 {table_counter}")
                    png_filename = self.render_html_table_as_image(str(table), table_counter, origin_number)
                    print(f"HTML 렌더링 결과: {png_filename}")
                    if png_filename is None:
                        # 실패시 fallback - 간단한 텍스트 이미지 생성
                        png_filename = f"Medical/Table/M_table_{origin_number}_{table_counter}.png"
                        try:
                            fig, ax = plt.subplots(figsize=(10, 6))
                            ax.text(0.5, 0.5, f'테이블 {table_counter}\n({len(df)} 행 x {len(df.columns)} 열)\n\n웹 렌더링 실패', 
                                   ha='center', va='center', fontsize=14, 
                                   bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray"))
                            ax.set_xlim(0, 1)
                            ax.set_ylim(0, 1)
                            ax.axis('off')
                            plt.tight_layout()
                            fig.savefig(png_filename, dpi=150, bbox_inches='tight')
                            plt.close(fig)
                        except Exception as e:
                            print(f"fallback 이미지 생성 실패: {e}")
                            png_filename = f"Medical/Table/M_table_{origin_number}_{table_counter}.png"

                    # 기본 메타 정보
                    table_entry = {
                        'table_number': table_counter,
                        'filename': png_filename,
                        'preview_text': ' | '.join(df.head(2).astype(str).fillna('').values.flatten()[:10]),
                        'rows': len(df),
                        'columns': len(df.columns),
                        'size': f"{len(df)}x{len(df.columns)}",
                        'image_size': None,
                        'position': f"panel[{p_idx}] table[{t_idx}]",
                        'extraction_method': 'html_panel_table_extraction'
                    }

                    table_info.append(table_entry)
                    table_counter += 1

            print(f"총 {len(table_info)}개의 테이블을 HTML에서 추출했습니다.")
            return table_info

        except Exception as e:
            print(f"HTML 테이블 추출 실패: {e}")
            return []
    
    def process_url(self, url, origin_number):
        """URL 처리 - PNG 저장 및 테이블 이미지 추출"""
        driver = None
        try:
            print(f"\n{'='*50}")
            print(f"처리 중: {url}")
            print(f"Origin Number: {origin_number}")
            print(f"{'='*50}")
            # 특정 사이트(단일 HTML에 모든 표가 숨겨진 경우)는 requests+BS4 방식으로 처리
            if 'davoshospital.co.kr' in url or 'page06_new.html' in url:
                print("특정 단일페이지 형식 감지 - HTML 직접 파싱으로 처리합니다.")
                table_info = self.extract_hidden_tables_from_url(url, origin_number)

                # 결과 정리 (간단한 메타)
                result = {
                    'origin_number': origin_number,
                    'url': url,
                    'page_title': url,
                    'png_filename': '',
                    'table_count': len(table_info),
                    'table_info': table_info,
                    'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'user_agent': 'N/A',
                    'window_size': 'N/A'
                }
                print(f"HTML 직접 파싱 처리 완료: {len(table_info)}개 테이블 추출")
                return result

            # WebDriver 설정
            driver = self.setup_webdriver()
            if not driver:
                return None
            
            # User-Agent 확인
            user_agent = driver.execute_script("return navigator.userAgent;")
            
            # 윈도우 크기 확인
            window_size = driver.get_window_size()
            
            # 웹페이지 로드
            print("웹페이지 로딩 중...")
            driver.get(url)
            
            # 페이지 제목 가져오기
            try:
                page_title = driver.title[:50] if driver.title else f"Page_{origin_number}"
                print(f"페이지 제목: {page_title}")
            except:
                page_title = f"Page_{origin_number}"
            
            # PNG 파일명 생성
            png_filename = f"Medical/Context/Origin/M_origin_{origin_number}.png"
            
            # PNG 저장
            png_success = self.save_page_as_png(driver, url, png_filename)
            if not png_success:
                return None
            
            # 테이블 이미지 캡처
            table_info = self.capture_tables_as_images(driver, origin_number)
            
            # 결과 정리
            result = {
                'origin_number': origin_number,
                'url': url,
                'page_title': page_title,
                'png_filename': png_filename,
                'table_count': len(table_info),
                'table_info': table_info,
                'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'user_agent': user_agent,
                'window_size': f"{window_size['width']}x{window_size['height']}"
            }
            
            print(f"URL 처리 완료: {len(table_info)}개 테이블 추출")
            return result
            
        except Exception as e:
            print(f"URL 처리 실패 ({url}): {e}")
            return None
            
        finally:
            if driver:
                driver.quit()
                print("WebDriver 종료")
    
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
                    'User Agent': result.get('user_agent', 'Unknown'),
                    'Window Size': result.get('window_size', 'Unknown')
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
                        'Rows': table['rows'],
                        'Columns': table['columns'],
                        'Preview Text': table['preview_text']
                    }
                    self.existing_data['table_data'].append(table_entry)
                
                # URL 집합 업데이트
                self.existing_data['existing_urls'].add(result['url'])
                
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
            table_dir = "/Users/gremh/tablemagnifier/Medical/Table"
            actual_file_count = 0
            try:
                table_files = [f for f in os.listdir(table_dir) if f.endswith('.png')]
                actual_file_count = len(table_files)
            except Exception as e:
                print(f"실제 파일 개수 확인 실패: {e}")
            
            # 결과 요약
            total_urls = len(self.existing_data['main_data'])
            total_tables_in_excel = len(self.existing_data['table_data'])
            
            print(f"\n{'='*60}")
            print(f"전체 데이터베이스 현황")
            print(f"{'='*60}")
            print(f"총 처리된 URL: {total_urls}개")
            print(f"엑셀에 기록된 테이블: {total_tables_in_excel}개")
            print(f"실제 저장된 파일: {actual_file_count}개")
            if total_tables_in_excel != actual_file_count:
                hidden_tables = total_tables_in_excel - actual_file_count
                print(f"숨겨진/건너뛴 테이블: {hidden_tables}개")
            print(f"최대 Origin Number: {self.existing_data['max_origin_number']}")
            print(f"엑셀 파일: {self.excel_filename}")
            print(f"PNG 저장 위치: Medical/Context/Origin/")
            print(f"테이블 이미지 저장 위치: Medical/Table/")
            print(f"{'='*60}")
            
        except Exception as e:
            print(f"엑셀 저장 실패: {e}")
    
    def run(self):
        """메인 실행 함수"""
        print("연속 PNG 및 테이블 이미지 추출 시작")
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # URL 읽기
        all_urls = self.read_urls()
        if not all_urls:
            print("처리할 URL이 없습니다.")
            return
        
        # 새로운 URL만 필터링
        new_urls = self.filter_new_urls(all_urls)
        
        if not new_urls:
            print("처리할 새로운 URL이 없습니다. 모든 URL이 이미 처리되었습니다.")
            
            # 테이블 디렉토리의 파일 개수 확인
            table_dir = "/Users/gremh/tablemagnifier/Medical/Table"
            try:
                table_files = [f for f in os.listdir(table_dir) if f.endswith('.png')]
                total_table_files = len(table_files)
                print(f"\n📁 테이블 디렉토리 파일 개수: {total_table_files}개")
                print(f"디렉토리 경로: {table_dir}")
            except Exception as e:
                print(f"테이블 디렉토리 파일 개수 확인 실패: {e}")
            
            return
        
        print(f"총 {len(new_urls)}개의 새로운 URL을 처리합니다.")
        
        # 새로운 결과 저장용
        new_results = []
        
        # 각 URL 처리
        for i, url in enumerate(new_urls):
            print(f"\n진행상황: {i+1}/{len(new_urls)}")
            
            # Origin Number 계산
            origin_number = self.get_next_origin_number()
            self.existing_data['max_origin_number'] = origin_number  # 즉시 업데이트
            
            result = self.process_url(url, origin_number)
            new_results.append(result)
            
            # 처리 결과를 즉시 엑셀에 저장 (중간 저장)
            if result:
                self.update_excel_data([result])
                self.save_to_excel()
                print(f"중간 저장 완료 (Origin {origin_number})")
            
            # 다음 URL 처리 전 잠시 대기
            if i < len(new_urls) - 1:
                print("다음 URL 처리를 위해 2초 대기...")
                time.sleep(2)
        
        # 최종 저장 (이미 중간에 저장되었지만 확인차 한 번 더)
        if any(new_results):
            print("최종 엑셀 파일 저장 확인...")
            self.save_to_excel()
        
        # 테이블 디렉토리의 파일 개수 확인
        table_dir = "/Users/gremh/tablemagnifier/Medical/Table"
        try:
            table_files = [f for f in os.listdir(table_dir) if f.endswith('.png')]
            total_table_files = len(table_files)
            print(f"\n📁 테이블 디렉토리 파일 개수: {total_table_files}개")
            print(f"디렉토리 경로: {table_dir}")
        except Exception as e:
            print(f"테이블 디렉토리 파일 개수 확인 실패: {e}")

        print(f"\n모든 작업이 완료되었습니다!")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    extractor = ContinuousPNGTableExtractor()
    extractor.run()