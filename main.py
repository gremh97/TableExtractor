#!/usr/bin/env python3
"""
Medical Table Extractor - Main Integration Script
URL과 PDF 파일 두 소스에서 테이블을 추출하는 통합 스크립트

실행 방법:
python main.py

기능:
1. URL 처리: continuous_table_extractor.py 실행
2. PDF 처리: pdf_processor.py 실행
3. 모든 결과를 Medical_Table_Results.xlsx에 통합
"""

import os
import sys
import subprocess
import time
from datetime import datetime

class MedicalTableExtractorMain:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.venv_python = os.path.join(self.base_dir, '.venv', 'bin', 'python')
        
        # 스크립트 경로
        self.url_processor = os.path.join(self.base_dir, 'continuous_table_extractor.py')
        self.pdf_processor = os.path.join(self.base_dir, 'pdf_processor_pdfplumber.py')
        
    def print_header(self):
        """시작 메시지 출력"""
        print("=" * 70)
        print("🏥 Medical Table Extractor - 통합 처리 시스템")
        print("=" * 70)
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"작업 디렉토리: {self.base_dir}")
        print()
        
    def check_files(self):
        """필요한 파일들이 존재하는지 확인"""
        print("📋 파일 존재 확인...")
        
        files_to_check = [
            ('가상환경 Python', self.venv_python),
            ('URL 처리기', self.url_processor),
            ('PDF 처리기', self.pdf_processor)
        ]
        
        missing_files = []
        for name, path in files_to_check:
            if os.path.exists(path):
                print(f"  ✅ {name}: {path}")
            else:
                print(f"  ❌ {name}: {path} (없음)")
                missing_files.append(name)
        
        if missing_files:
            print(f"\n❌ 다음 파일들이 없습니다: {', '.join(missing_files)}")
            return False
        
        print("✅ 모든 필수 파일이 존재합니다.\n")
        return True
        
    def check_source_files(self):
        """처리할 소스 파일들 확인"""
        print("📁 소스 파일 확인...")
        
        # URL 파일 확인
        url_file = os.path.join(self.base_dir, 'urls.txt')
        url_count = 0
        if os.path.exists(url_file):
            try:
                with open(url_file, 'r', encoding='utf-8') as f:
                    urls = [line.strip() for line in f if line.strip()]
                    url_count = len(urls)
                print(f"  📄 URL 파일: {url_count}개 URL 발견")
            except Exception as e:
                print(f"  ⚠️  URL 파일 읽기 오류: {e}")
        else:
            print(f"  ⚠️  URL 파일 없음: {url_file}")
        
        # PDF 파일 확인
        pdf_dir = os.path.join(self.base_dir, 'temperal_pdf')
        pdf_count = 0
        if os.path.exists(pdf_dir):
            pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]
            pdf_count = len(pdf_files)
            print(f"  📑 PDF 파일: {pdf_count}개 PDF 발견")
            for pdf in pdf_files[:5]:  # 처음 5개만 표시
                print(f"    - {pdf}")
            if pdf_count > 5:
                print(f"    ... 및 {pdf_count - 5}개 더")
        else:
            print(f"  ⚠️  PDF 디렉토리 없음: {pdf_dir}")
        
        total_sources = url_count + pdf_count
        print(f"  📊 총 처리 예정: URL {url_count}개 + PDF {pdf_count}개 = {total_sources}개")
        
        if total_sources == 0:
            print("  ⚠️  처리할 소스가 없습니다. URL 파일이나 PDF 파일을 확인해주세요.")
            return False
        
        print()
        return True
        
    def run_script(self, script_path, script_name):
        """스크립트 실행"""
        print(f"🚀 {script_name} 실행 중...")
        print(f"   명령어: {self.venv_python} {script_path}")
        print("-" * 50)
        
        try:
            # subprocess로 스크립트 실행
            result = subprocess.run(
                [self.venv_python, script_path],
                cwd=self.base_dir,
                capture_output=False,  # 실시간 출력을 위해 False
                text=True,
                check=True
            )
            
            print("-" * 50)
            print(f"✅ {script_name} 완료!")
            print()
            return True
            
        except subprocess.CalledProcessError as e:
            print("-" * 50)
            print(f"❌ {script_name} 실행 실패!")
            print(f"   오류 코드: {e.returncode}")
            if e.stdout:
                print(f"   출력: {e.stdout}")
            if e.stderr:
                print(f"   에러: {e.stderr}")
            print()
            return False
            
        except Exception as e:
            print("-" * 50)
            print(f"❌ {script_name} 실행 중 예외 발생!")
            print(f"   오류: {e}")
            print()
            return False
    
    def show_final_status(self):
        """최종 상태 표시"""
        print("=" * 70)
        print("📊 처리 완료 - 최종 상태")
        print("=" * 70)
        
        # Excel 파일 확인
        excel_file = os.path.join(self.base_dir, 'Medical_Table_Results.xlsx')
        if os.path.exists(excel_file):
            try:
                import pandas as pd
                main_df = pd.read_excel(excel_file, sheet_name='Main Results')
                table_df = pd.read_excel(excel_file, sheet_name='Table Details')
                
                print(f"📋 Excel 파일: {excel_file}")
                print(f"   📄 총 처리된 항목: {len(main_df)}개")
                print(f"   🖼️  추출된 테이블: {len(table_df)}개")
                
                # URL vs PDF 분류
                url_count = len(main_df[~main_df['URL'].str.startswith('PDF_FILE:')])
                pdf_count = len(main_df[main_df['URL'].str.startswith('PDF_FILE:')])
                print(f"   🌐 URL 처리: {url_count}개")
                print(f"   📑 PDF 처리: {pdf_count}개")
                
            except Exception as e:
                print(f"   ⚠️  Excel 파일 분석 실패: {e}")
        else:
            print("   ❌ Excel 파일이 생성되지 않았습니다.")
        
        # 디렉토리 상태
        origin_dir = os.path.join(self.base_dir, 'Medical', 'Context', 'Origin')
        table_dir = os.path.join(self.base_dir, 'Medical', 'Table')
        
        if os.path.exists(origin_dir):
            origin_files = len([f for f in os.listdir(origin_dir) if f.startswith('M_origin_')])
            print(f"📁 원본 파일: {origin_files}개 저장")
        
        if os.path.exists(table_dir):
            table_files = len([f for f in os.listdir(table_dir) if f.startswith('M_table_')])
            print(f"🖼️  테이블 이미지: {table_files}개 저장")
        
        print(f"\n🎉 모든 처리가 완료되었습니다!")
        print(f"완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)
    
    def run(self):
        """메인 실행 함수"""
        # 시작 메시지
        self.print_header()
        
        # 파일 존재 확인
        if not self.check_files():
            print("❌ 필수 파일이 없어서 실행을 중단합니다.")
            return False
        
        # 소스 파일 확인
        if not self.check_source_files():
            print("❌ 처리할 소스가 없어서 실행을 중단합니다.")
            return False
        
        # 사용자 확인
        print("🔄 처리를 시작하시겠습니까? (y/N): ", end="")
        try:
            user_input = input().strip().lower()
            if user_input not in ['y', 'yes', '예', 'ㅇ']:
                print("처리가 취소되었습니다.")
                return False
        except KeyboardInterrupt:
            print("\n처리가 취소되었습니다.")
            return False
        
        print()
        
        # 1단계: URL 처리
        print("=" * 50)
        print("1단계: URL에서 테이블 추출")
        print("=" * 50)
        
        url_success = self.run_script(self.url_processor, "URL 테이블 추출기")
        
        if url_success:
            print("⏳ URL 처리와 PDF 처리 사이에 2초 대기...")
            time.sleep(2)
        
        # 2단계: PDF 처리
        print("=" * 50)
        print("2단계: PDF에서 테이블 추출")
        print("=" * 50)
        
        pdf_success = self.run_script(self.pdf_processor, "PDF 테이블 추출기")
        
        # 최종 결과
        print()
        self.show_final_status()
        
        # 성공 여부 반환
        return url_success and pdf_success

def main():
    """프로그램 진입점"""
    try:
        extractor = MedicalTableExtractorMain()
        success = extractor.run()
        
        # 종료 코드 설정
        sys.exit(0 if success else 1)
        
    except KeyboardInterrupt:
        print("\n\n🛑 사용자가 프로그램을 중단했습니다.")
        sys.exit(1)
        
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류가 발생했습니다: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()